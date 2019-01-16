package com.zjf.activity;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.util.Log;
import android.util.Xml;

import com.zjf.util.PictureData;

public class FR {
    private String nameStr;
    public Range range = null;
    public HWPFDocument hwpf = null;
    public String htmlPath;
    public String picturePath;
    public List pictures;
    public TableIterator tableIterator;
    public int presentPicture = 0;
    public int screenWidth;
    public FileOutputStream output;
    public File myFile;
    StringBuffer lsb = new StringBuffer();
    String returnPath = "";
    static final int BUFFER = 2048;

    public FR(String namepath) {
        // this.screenWidth =
        // this.getWindowManager().getDefaultDisplay().getWidth() -
        // 10;//���ÿ��Ϊ��Ļ���-10
        this.nameStr = namepath;
        read();
    }

    public void read() {

        if (this.nameStr.endsWith(".doc")) {
            this.getRange();
            this.makeFile();
            this.readDOC();
            returnPath = "file:///" + this.htmlPath;
            // this.view.loadUrl("file:///" + this.htmlPath);
            System.out.println("htmlPath" + this.htmlPath);
        }
        if (this.nameStr.endsWith(".docx")) {
            this.makeFile();
            this.readDOCX();
            returnPath = "file:///" + this.htmlPath;
            // this.view.loadUrl("file:///" + this.htmlPath);
            System.out.println("htmlPath" + this.htmlPath);
        }
        if (this.nameStr.endsWith(".xls")) {

            try {
                this.makeFile();
                this.readXLS();
                returnPath = "file:///" + this.htmlPath;
                // this.view.loadUrl("file:///" + this.htmlPath);
                System.out.println("htmlPath" + this.htmlPath);
            } catch (Exception e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }

        }
        if (this.nameStr.endsWith(".xlsx")) {
            this.makeFile();
            try {
                this.readXLSX();
            } catch (FileNotFoundException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            returnPath = "file:///" + this.htmlPath;
            // this.view.loadUrl("file:///" + this.htmlPath);
            System.out.println("htmlPath" + this.htmlPath);
        }

        if (this.nameStr.endsWith(".pptx")) {
            this.makeFile();
            this.readPPTX();
            returnPath = "file:///" + this.htmlPath;
            System.out.println("pptxhtmlPath=====" + this.htmlPath);
        }

    }

    /* ��ȡword�е�����д��sdcard�ϵ�.html�ļ��� */
    public void readDOC() {

        try {
            myFile = new File(htmlPath);
            output = new FileOutputStream(myFile);
            String head = "<html><meta charset=\"utf-8\"><body>";
            String tagBegin = "<p>";
            String tagEnd = "</p>";
            output.write(head.getBytes());
            int numParagraphs = range.numParagraphs();// �õ�ҳ�����еĶ�����
            for (int i = 0; i < numParagraphs; i++) { // ����������
                Paragraph p = range.getParagraph(i); // �õ��ĵ��е�ÿһ������
                if (p.isInTable()) {
                    int temp = i;
                    if (tableIterator.hasNext()) {
                        String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
                        String tableEnd = "</table>";
                        String rowBegin = "<tr>";
                        String rowEnd = "</tr>";
                        String colBegin = "<td>";
                        String colEnd = "</td>";
                        Table table = tableIterator.next();
                        output.write(tableBegin.getBytes());
                        int rows = table.numRows();
                        for (int r = 0; r < rows; r++) {
                            output.write(rowBegin.getBytes());
                            TableRow row = table.getRow(r);
                            int cols = row.numCells();
                            int rowNumParagraphs = row.numParagraphs();
                            int colsNumParagraphs = 0;
                            for (int c = 0; c < cols; c++) {
                                output.write(colBegin.getBytes());
                                TableCell cell = row.getCell(c);
                                int max = temp + cell.numParagraphs();
                                colsNumParagraphs = colsNumParagraphs
                                        + cell.numParagraphs();
                                for (int cp = temp; cp < max; cp++) {
                                    Paragraph p1 = range.getParagraph(cp);
                                    output.write(tagBegin.getBytes());
                                    writeParagraphContent(p1);
                                    output.write(tagEnd.getBytes());
                                    temp++;
                                }
                                output.write(colEnd.getBytes());
                            }
                            int max1 = temp + rowNumParagraphs;
                            for (int m = temp + colsNumParagraphs; m < max1; m++) {
                                temp++;
                            }
                            output.write(rowEnd.getBytes());
                        }
                        output.write(tableEnd.getBytes());
                    }
                    i = temp;
                } else {
                    output.write(tagBegin.getBytes());
                    writeParagraphContent(p);
                    output.write(tagEnd.getBytes());
                }
            }
            String end = "</body></html>";
            output.write(end.getBytes());
            output.close();
        } catch (Exception e) {

            System.out.println("readAndWrite Exception:" + e.getMessage());
            e.printStackTrace();
        }
    }

    public void readDOCX() {
        String river = "";
        try {
            this.myFile = new File(this.htmlPath);// newһ��File,·��Ϊhtml�ļ�
            this.output = new FileOutputStream(this.myFile);// newһ����,Ŀ��Ϊhtml�ļ�
            String head = "<!DOCTYPE><html><meta charset=\"utf-8\"><body>";// ����ͷ�ļ�,�����������utf-8,��Ȼ���������
            String end = "</body></html>";
            String tagBegin = "<p>";// ���俪ʼ,��ǿ�ʼ?
            String tagEnd = "</p>";// �������
            String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
            String tableEnd = "</table>";
            String rowBegin = "<tr>";
            String rowEnd = "</tr>";
            String colBegin = "<td>";
            String colEnd = "</td>";
            String style = "style=\"";
            this.output.write(head.getBytes());// д��ͷ��
            ZipFile xlsxFile = new ZipFile(new File(this.nameStr));
            ZipEntry sharedStringXML = xlsxFile.getEntry("word/document.xml");
            InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
            XmlPullParser xmlParser = Xml.newPullParser();
            xmlParser.setInput(inputStream, "utf-8");
            int evtType = xmlParser.getEventType();
            boolean isTable = false; // �Ǳ�� ����ͳ�� �� �� ��
            boolean isSize = false; // ��С״̬
            boolean isColor = false; // ��ɫ״̬
            boolean isCenter = false; // ����״̬
            boolean isRight = false; // ����״̬
            boolean isItalic = false; // ��б��
            boolean isUnderline = false; // ���»���
            boolean isBold = false; // �Ӵ�
            boolean isR = false; // ���Ǹ�r��
            boolean isStyle = false;
            int pictureIndex = 1; // docx ѹ�����е�ͼƬ�� iamge1 ��ʼ ����������1��ʼ
            while (evtType != XmlPullParser.END_DOCUMENT) {
                switch (evtType) {

                // ��ʼ��ǩ
                case XmlPullParser.START_TAG:
                    String tag = xmlParser.getName();

                    if (tag.equalsIgnoreCase("r")) {
                        isR = true;
                    }
                    if (tag.equalsIgnoreCase("u")) { // �ж��»���
                        isUnderline = true;
                    }
                    if (tag.equalsIgnoreCase("jc")) { // �ж϶��뷽ʽ
                        String align = xmlParser.getAttributeValue(0);
                        if (align.equals("center")) {
                            this.output.write("<center>".getBytes());
                            isCenter = true;
                        }
                        if (align.equals("right")) {
                            this.output.write("<div align=\"right\">"
                                    .getBytes());
                            isRight = true;
                        }
                    }

                    if (tag.equalsIgnoreCase("color")) { // �ж���ɫ

                        String color = xmlParser.getAttributeValue(0);

                        this.output
                                .write(("<span style=\"color:" + color + ";\">")
                                        .getBytes());
                        isColor = true;
                    }
                    if (tag.equalsIgnoreCase("sz")) { // �жϴ�С
                        if (isR == true) {
                            int size = decideSize(Integer.valueOf(xmlParser
                                    .getAttributeValue(0)));
                            this.output.write(("<font size=" + size + ">")
                                    .getBytes());
                            isSize = true;
                        }
                    }
                    // �����Ǳ����
                    if (tag.equalsIgnoreCase("tbl")) { // ��⵽tbl ���ʼ
                        this.output.write(tableBegin.getBytes());
                        isTable = true;
                    }
                    if (tag.equalsIgnoreCase("tr")) { // ��
                        this.output.write(rowBegin.getBytes());
                    }
                    if (tag.equalsIgnoreCase("tc")) { // ��
                        this.output.write(colBegin.getBytes());
                    }

                    if (tag.equalsIgnoreCase("pic")) { // ��⵽��ǩ pic ͼƬ
                        String entryName_jpeg = "word/media/image"
                                + pictureIndex + ".jpeg";
                        String entryName_png = "word/media/image"
                                + pictureIndex + ".png";
                        String entryName_gif = "word/media/image"
                                + pictureIndex + ".gif";
                        String entryName_wmf = "word/media/image"
                                + pictureIndex + ".wmf";
                        ZipEntry sharePicture = null;
                        InputStream pictIS = null;
                        sharePicture = xlsxFile.getEntry(entryName_jpeg);
                        // һ��Ϊ��ȡdocx��ͼƬ ת��Ϊ������
                        if (sharePicture == null) {
                            sharePicture = xlsxFile.getEntry(entryName_png);
                        }
                        if (sharePicture == null) {
                            sharePicture = xlsxFile.getEntry(entryName_gif);
                        }
                        if (sharePicture == null) {
                            sharePicture = xlsxFile.getEntry(entryName_wmf);
                        }

                        if (sharePicture != null) {
                            pictIS = xlsxFile.getInputStream(sharePicture);
                            ByteArrayOutputStream pOut = new ByteArrayOutputStream();
                            byte[] bt = null;
                            byte[] b = new byte[1000];
                            int len = 0;
                            while ((len = pictIS.read(b)) != -1) {
                                pOut.write(b, 0, len);
                            }
                            pictIS.close();
                            pOut.close();
                            bt = pOut.toByteArray();
                            Log.i("byteArray", "" + bt);
                            if (pictIS != null)
                                pictIS.close();
                            if (pOut != null)
                                pOut.close();
                            writeDOCXPicture(bt);
                        }

                        pictureIndex++; // ת��һ�ź� ����+1
                    }

                    if (tag.equalsIgnoreCase("b")) { // ��⵽�Ӵֱ�ǩ
                        isBold = true;
                    }
                    if (tag.equalsIgnoreCase("p")) {// ��⵽ p ��ǩ
                        if (isTable == false) { // ����ڱ���� ������
                            this.output.write(tagBegin.getBytes());
                        }
                    }
                    if (tag.equalsIgnoreCase("i")) { // б��
                        isItalic = true;
                    }
                    // ��⵽ֵ ��ǩ
                    if (tag.equalsIgnoreCase("t")) {
                        if (isBold == true) { // �Ӵ�
                            this.output.write("<b>".getBytes());
                        }
                        if (isUnderline == true) { // ��⵽�»��߱�ǩ,����<u>
                            this.output.write("<u>".getBytes());
                        }
                        if (isItalic == true) { // ��⵽б���ǩ,����<i>
                            output.write("<i>".getBytes());
                        }
                        river = xmlParser.nextText();
                        this.output.write(river.getBytes()); // д����ֵ
                        if (isItalic == true) { // ��⵽б���ǩ,������ֵ֮��,����</i>,����б��״̬=false
                            this.output.write("</i>".getBytes());
                            isItalic = false;
                        }
                        if (isUnderline == true) {// ��⵽�»��߱�ǩ,������ֵ֮��,����</u>,�����»���״̬=false
                            this.output.write("</u>".getBytes());
                            isUnderline = false;
                        }
                        if (isBold == true) { // �Ӵ�
                            this.output.write("</b>".getBytes());
                            isBold = false;
                        }
                        if (isSize == true) { // ��⵽��С����,���������ǩ
                            this.output.write("</font>".getBytes());
                            isSize = false;
                        }
                        if (isColor == true) { // ��⵽��ɫ���ô���,���������ǩ
                            this.output.write("</span>".getBytes());
                            isColor = false;
                        }
                        if (isCenter == true) { // ��⵽����,���������ǩ
                            this.output.write("</center>".getBytes());
                            isCenter = false;
                        }
                        if (isRight == true) { // ���Ҳ���ʹ��<right></right>,ʹ��div���ܻ���״��,������
                            this.output.write("</div>".getBytes());
                            isRight = false;
                        }
                    }
                    break;
                // ������ǩ
                case XmlPullParser.END_TAG:
                    String tag2 = xmlParser.getName();
                    if (tag2.equalsIgnoreCase("tbl")) { // ��⵽������,���ı��״̬
                        this.output.write(tableEnd.getBytes());
                        isTable = false;
                    }
                    if (tag2.equalsIgnoreCase("tr")) { // �н���
                        this.output.write(rowEnd.getBytes());
                    }
                    if (tag2.equalsIgnoreCase("tc")) { // �н���
                        this.output.write(colEnd.getBytes());
                    }
                    if (tag2.equalsIgnoreCase("p")) { // p����,����ڱ���о�����
                        if (isTable == false) {
                            this.output.write(tagEnd.getBytes());
                        }
                    }
                    if (tag2.equalsIgnoreCase("r")) {
                        isR = false;
                    }
                    break;
                default:
                    break;
                }
                evtType = xmlParser.next();
            }
            this.output.write(end.getBytes());
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlPullParserException e) {
            e.printStackTrace();
        }
        if (river == null) {
            river = "�����ļ���������";
        }
    }

    public StringBuffer readXLS() throws Exception {

        myFile = new File(htmlPath);
        output = new FileOutputStream(myFile);
        lsb.append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>");
        lsb.append("<head><meta http-equiv=Content-Type content='text/html; charset=utf-8'><meta name=ProgId content=Excel.Sheet>");
        HSSFSheet sheet = null;

        String excelFileName = nameStr;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(
                    excelFileName)); // ������Excel

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                sheet = workbook.getSheetAt(sheetIndex);// �����е�sheet
                String sheetName = workbook.getSheetName(sheetIndex); // sheetName
                if (workbook.getSheetAt(sheetIndex) != null) {
                    sheet = workbook.getSheetAt(sheetIndex);// ��ò�Ϊ�յ����sheet
                    if (sheet != null) {
                        int firstRowNum = sheet.getFirstRowNum(); // ��һ��
                        int lastRowNum = sheet.getLastRowNum(); // ���һ��
                        // ����Table
                        lsb.append("<table width=\"100%\" style=\"border:1px solid #000;border-width:1px 0 0 1px;margin:2px 0 2px 0;border-collapse:collapse;\">");
                        for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
                            if (sheet.getRow(rowNum) != null) {// ����в�Ϊ�գ�
                                HSSFRow row = sheet.getRow(rowNum);
                                short firstCellNum = row.getFirstCellNum(); // ���еĵ�һ����Ԫ��
                                short lastCellNum = row.getLastCellNum(); // ���е����һ����Ԫ��
                                int height = (int) (row.getHeight() / 15.625); // �еĸ߶�
                                lsb.append("<tr height=\""
                                        + height
                                        + "\" style=\"border:1px solid #000;border-width:0 1px 1px 0;margin:2px 0 2px 0;\">");
                                for (short cellNum = firstCellNum; cellNum <= lastCellNum; cellNum++) { // ѭ�����е�ÿһ����Ԫ��
                                    HSSFCell cell = row.getCell(cellNum);
                                    if (cell != null) {
                                        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                                            continue;
                                        } else {
                                            StringBuffer tdStyle = new StringBuffer(
                                                    "<td style=\"border:1px solid #000; border-width:0 1px 1px 0;margin:2px 0 2px 0; ");
                                            HSSFCellStyle cellStyle = cell
                                                    .getCellStyle();
                                            HSSFPalette palette = workbook
                                                    .getCustomPalette(); // ��HSSFPalette��������ɫ�Ĺ��ʱ�׼��ʽ
                                            HSSFColor hColor = palette
                                                    .getColor(cellStyle
                                                            .getFillForegroundColor());
                                            HSSFColor hColor2 = palette
                                                    .getColor(cellStyle
                                                            .getFont(workbook)
                                                            .getColor());

                                            String bgColor = convertToStardColor(hColor);// ������ɫ
                                            short boldWeight = cellStyle
                                                    .getFont(workbook)
                                                    .getBoldweight(); // �����ϸ
                                            short fontHeight = (short) (cellStyle
                                                    .getFont(workbook)
                                                    .getFontHeight() / 2); // �����С
                                            String fontColor = convertToStardColor(hColor2); // ������ɫ
                                            if (bgColor != null
                                                    && !"".equals(bgColor
                                                            .trim())) {
                                                tdStyle.append(" background-color:"
                                                        + bgColor + "; ");
                                            }
                                            if (fontColor != null
                                                    && !"".equals(fontColor
                                                            .trim())) {
                                                tdStyle.append(" color:"
                                                        + fontColor + "; ");
                                            }
                                            tdStyle.append(" font-weight:"
                                                    + boldWeight + "; ");
                                            tdStyle.append(" font-size: "
                                                    + fontHeight + "%;");
                                            lsb.append(tdStyle + "\"");

                                            int width = (int) (sheet
                                                    .getColumnWidth(cellNum) / 35.7); //
                                            int cellReginCol = getMergerCellRegionCol(
                                                    sheet, rowNum, cellNum); // �ϲ����У�solspan��
                                            int cellReginRow = getMergerCellRegionRow(
                                                    sheet, rowNum, cellNum);// �ϲ����У�rowspan��
                                            String align = convertAlignToHtml(cellStyle
                                                    .getAlignment()); //
                                            String vAlign = convertVerticalAlignToHtml(cellStyle
                                                    .getVerticalAlignment());

                                            lsb.append(" align=\"" + align
                                                    + "\" valign=\"" + vAlign
                                                    + "\" width=\"" + width
                                                    + "\" ");

                                            lsb.append(" colspan=\""
                                                    + cellReginCol
                                                    + "\" rowspan=\""
                                                    + cellReginRow + "\"");
                                            lsb.append(">" + getCellValue(cell)
                                                    + "</td>");
                                        }
                                    }
                                }
                                lsb.append("</tr>");
                            }
                        }
                    }

                }

            }
            output.write(lsb.toString().getBytes());
        } catch (FileNotFoundException e) {
            throw new Exception("�ļ� " + excelFileName + " û���ҵ�!");
        } catch (IOException e) {
            throw new Exception("�ļ� " + excelFileName + " �������("
                    + e.getMessage() + ")!");
        }
        return lsb;
    }

    public void readXLSX() throws FileNotFoundException{
        myFile = new File(htmlPath);
        output = new FileOutputStream(myFile);
        lsb.append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>");
        lsb.append("<head><meta http-equiv=Content-Type content='text/html; charset=utf-8'><meta name=ProgId content=Excel.Sheet>");
        XSSFSheet sheet = null;

        String excelFileName = nameStr;
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(
                    excelFileName)); // ������Excel

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                sheet = workbook.getSheetAt(sheetIndex);// �����е�sheet
                Log.e("=================", sheet.getSheetName());
            }
        }catch(Exception e){
            e.printStackTrace();
            
        }
    }
    
//    public void readXLSX() {
//        try {
//            this.myFile = new File(this.htmlPath);// newһ��File,·��Ϊhtml�ļ�
//            this.output = new FileOutputStream(this.myFile);// newһ����,Ŀ��Ϊhtml�ļ�
//            String head = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\"\"http://www.w3.org/TR/html4/loose.dtd\"><html><meta charset=\"utf-8\"><head></head><body>";// ����ͷ�ļ�,�����������utf-8,��Ȼ���������
//            String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
//            String tableEnd = "</table>";
//            String rowBegin = "<tr>";
//            String rowEnd = "</tr>";
//            String colBegin = "<td>";
//            String colEnd = "</td>";
//            String end = "</body></html>";
//            this.output.write(head.getBytes());
//            this.output.write(tableBegin.getBytes());
//            String str = "";
//            String v = null;
//            boolean flat = false;
//            List<String> ls = new ArrayList<String>();
//            try {
//                ZipFile xlsxFile = new ZipFile(new File(this.nameStr));// ��ַ
//                ZipEntry sharedStringXML = xlsxFile
//                        .getEntry("xl/sharedStrings.xml");// �����ַ���
//                InputStream inputStream = xlsxFile
//                        .getInputStream(sharedStringXML);// ������ Ŀ������Ĺ����ַ���
//                XmlPullParser xmlParser = Xml.newPullParser();// new ������
//                xmlParser.setInput(inputStream, "utf-8");// ���ý���������
//                int evtType = xmlParser.getEventType();// ��ȡ���������¼�����
//                while (evtType != XmlPullParser.END_DOCUMENT) {// ��������� �ĵ�����
//                    switch (evtType) {
//                    case XmlPullParser.START_TAG: // ��ǩ��ʼ
//                        String tag = xmlParser.getName();
//                        if (tag.equalsIgnoreCase("t")) {
//                            ls.add(xmlParser.nextText());
//                        }
//                        break;
//                    case XmlPullParser.END_TAG: // ��ǩ����
//                        break;
//                    default:
//                        break;
//                    }
//                    evtType = xmlParser.next();
//                }
//                ZipEntry sheetXML = xlsxFile
//                        .getEntry("xl/worksheets/sheet1.xml");
//                InputStream inputStreamsheet = xlsxFile
//                        .getInputStream(sheetXML);
//                XmlPullParser xmlParsersheet = Xml.newPullParser();
//                xmlParsersheet.setInput(inputStreamsheet, "utf-8");
//                int evtTypesheet = xmlParsersheet.getEventType();
//                this.output.write(rowBegin.getBytes());
//                int i = -1;
//                while (evtTypesheet != XmlPullParser.END_DOCUMENT) {
//                    switch (evtTypesheet) {
//                    case XmlPullParser.START_TAG: // ��ǩ��ʼ
//                        String tag = xmlParsersheet.getName();
//                        if (tag.equalsIgnoreCase("row")) {
//                        } else {
//                            if (tag.equalsIgnoreCase("c")) {
//                                String t = xmlParsersheet.getAttributeValue(
//                                        null, "t");
//                                if (t != null) {
//                                    flat = true;
//                                    System.out.println(flat + "��");
//                                } else {// û������ʱ ������n��,�����ҵ��� ����<td></td> ��ʾ�ո�
//                                    this.output.write(colBegin.getBytes());
//                                    this.output.write(colEnd.getBytes());
//                                    System.out.println(flat + "û��");
//                                    flat = false;
//                                }
//                            } else {
//                                if (tag.equalsIgnoreCase("v")) {
//                                    v = xmlParsersheet.nextText();
//                                    this.output.write(colBegin.getBytes());
//                                    if (v != null) {
//                                        if (flat) {
//                                            str = ls.get(Integer.parseInt(v));
//                                        } else {
//                                            str = v;
//                                        }
//                                        this.output.write(str.getBytes());
//                                        this.output.write(colEnd.getBytes());
//                                    }
//                                }
//                            }
//                        }
//                        break;
//                    case XmlPullParser.END_TAG:
//                        if (xmlParsersheet.getName().equalsIgnoreCase("row")
//                                && v != null) {
//                            if (i == 1) {
//                                this.output.write(rowEnd.getBytes());
//                                this.output.write(rowBegin.getBytes());
//                                i = 1;
//                            } else {
//                                this.output.write(rowBegin.getBytes());
//                            }
//                        }
//                        break;
//                    }
//                    evtTypesheet = xmlParsersheet.next();
//                }
//                System.out.println(str);
//            } catch (ZipException e) {
//                e.printStackTrace();
//            } catch (IOException e) {
//                e.printStackTrace();
//            } catch (XmlPullParserException e) {
//                e.printStackTrace();
//            }
//            if (str == null) {
//                str = "�����ļ���������";
//            }
//            this.output.write(rowEnd.getBytes());
//            this.output.write(tableEnd.getBytes());
//            this.output.write(end.getBytes());
//        } catch (Exception e) {
//            System.out.println("readAndWrite Exception");
//        }
//    }

    public void readPPTX() {
        List<String> ls = new ArrayList<String>();
        String river = "";
        ZipFile xlsxFile = null;
        try {
            xlsxFile = new ZipFile(new File(this.nameStr));// pptx���ն�ȡzip��ʽ��ȡ
        } catch (ZipException e1) {
            e1.printStackTrace();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        try {
            ZipEntry sharedStringXML = xlsxFile.getEntry("[Content_Types].xml");// �ҵ����������ݵ��ļ�
            InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);// ���õ��ļ���
            XmlPullParser xmlParser = Xml.newPullParser();// ʵ����pull
            xmlParser.setInput(inputStream, "utf-8");// �����Ž�pull��
            int evtType = xmlParser.getEventType();// �õ���ǩ���͵�״̬
            while (evtType != XmlPullParser.END_DOCUMENT) {// ѭ����ȡ��
                switch (evtType) {
                case XmlPullParser.START_TAG: // �жϱ�ǩ��ʼ��ȡ
                    String tag = xmlParser.getName();// �õ���ǩ
                    if (tag.equalsIgnoreCase("Override")) {
                        String s = xmlParser
                                .getAttributeValue(null, "PartName");
                        if (s.lastIndexOf("/ppt/slides/slide") == 0) {
                            ls.add(s);
                        }
                    }
                    break;
                case XmlPullParser.END_TAG:// ��ǩ��ȡ����
                    break;
                default:
                    break;
                }
                evtType = xmlParser.next();// ��ȡ��һ����ǩ
            }
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlPullParserException e) {
            e.printStackTrace();
        }
        for (int i = 1; i < (ls.size() + 1); i++) {// ������6�Żõ�Ƭ
            river += "��" + i + "�š�������������������������������" + "\n";
            try {
                ZipEntry sharedStringXML = xlsxFile.getEntry("ppt/slides/slide"
                        + i + ".xml");// �ҵ����������ݵ��ļ�
                InputStream inputStream = xlsxFile
                        .getInputStream(sharedStringXML);// ���õ��ļ���
                XmlPullParser xmlParser = Xml.newPullParser();// ʵ����pull
                xmlParser.setInput(inputStream, "utf-8");// �����Ž�pull��
                int evtType = xmlParser.getEventType();// �õ���ǩ���͵�״̬
                while (evtType != XmlPullParser.END_DOCUMENT) {// ѭ����ȡ��
                    switch (evtType) {
                    case XmlPullParser.START_TAG: // �жϱ�ǩ��ʼ��ȡ
                        String tag = xmlParser.getName();// �õ���ǩ
                        if (tag.equalsIgnoreCase("t")) {
                            river += xmlParser.nextText() + "\n";
                        }
                        break;
                    case XmlPullParser.END_TAG:// ��ǩ��ȡ����
                        break;
                    default:
                        break;
                    }
                    evtType = xmlParser.next();// ��ȡ��һ����ǩ
                }
            } catch (ZipException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (XmlPullParserException e) {
                e.printStackTrace();
            }
        }
        if (river == null) {
            river = "�����ļ���������";
        }
    }

    /**
     * ȡ�õ�Ԫ���ֵ
     * 
     * @param cell
     * @return
     * @throws IOException
     */
    private static Object getCellValue(HSSFCell cell) throws IOException {
        Object value = "";
        if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
            value = cell.getRichStringCellValue().toString();
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                Date date = (Date) cell.getDateCellValue();
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                value = sdf.format(date);
            } else {
                double value_temp = (double) cell.getNumericCellValue();
                BigDecimal bd = new BigDecimal(value_temp);
                BigDecimal bd1 = bd.setScale(3, bd.ROUND_HALF_UP);
                value = bd1.doubleValue();

                DecimalFormat format = new DecimalFormat("#0.###");
                value = format.format(cell.getNumericCellValue());

            }
        }
        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
            value = "";
        }
        return value;
    }

    /**
     * �жϵ�Ԫ���ڲ��ںϲ���Ԫ��Χ�ڣ�����ǣ���ȡ��ϲ���������
     * 
     * @param sheet
     *            ������
     * @param cellRow
     *            ���жϵĵ�Ԫ����к�
     * @param cellCol
     *            ���жϵĵ�Ԫ����к�
     * @return
     * @throws IOException
     */
    private static int getMergerCellRegionCol(HSSFSheet sheet, int cellRow,
            int cellCol) throws IOException {
        int retVal = 0;
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
            int firstRow = cra.getFirstRow(); // �ϲ���Ԫ��CELL��ʼ��
            int firstCol = cra.getFirstColumn(); // �ϲ���Ԫ��CELL��ʼ��
            int lastRow = cra.getLastRow(); // �ϲ���Ԫ��CELL������
            int lastCol = cra.getLastColumn(); // �ϲ���Ԫ��CELL������
            if (cellRow >= firstRow && cellRow <= lastRow) { // �жϸõ�Ԫ���Ƿ����ںϲ���Ԫ����
                if (cellCol >= firstCol && cellCol <= lastCol) {
                    retVal = lastCol - firstCol + 1; // �õ��ϲ�������
                    break;
                }
            }
        }
        return retVal;
    }

    /**
     * �жϵ�Ԫ���Ƿ��Ǻϲ��ĵ�������ǣ���ȡ��ϲ���������
     * 
     * @param sheet
     *            ��
     * @param cellRow
     *            ���жϵĵ�Ԫ����к�
     * @param cellCol
     *            ���жϵĵ�Ԫ����к�
     * @return
     * @throws IOException
     */
    private static int getMergerCellRegionRow(HSSFSheet sheet, int cellRow,
            int cellCol) throws IOException {
        int retVal = 0;
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
            int firstRow = cra.getFirstRow(); // �ϲ���Ԫ��CELL��ʼ��
            int firstCol = cra.getFirstColumn(); // �ϲ���Ԫ��CELL��ʼ��
            int lastRow = cra.getLastRow(); // �ϲ���Ԫ��CELL������
            int lastCol = cra.getLastColumn(); // �ϲ���Ԫ��CELL������
            if (cellRow >= firstRow && cellRow <= lastRow) { // �жϸõ�Ԫ���Ƿ����ںϲ���Ԫ����
                if (cellCol >= firstCol && cellCol <= lastCol) {
                    retVal = lastRow - firstRow + 1; // �õ��ϲ�������
                    break;
                }
            }
        }
        return 0;
    }

    /**
     * ��Ԫ�񱳾�ɫת��
     * 
     * @param hc
     * @return
     */
    private String convertToStardColor(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            int a = HSSFColor.AUTOMATIC.index;
            int b = hc.getIndex();
            if (a == b) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                String str;
                String str_tmp = Integer.toHexString(hc.getTriplet()[i]);
                if (str_tmp != null && str_tmp.length() < 2) {
                    str = "0" + str_tmp;
                } else {
                    str = str_tmp;
                }
                sb.append(str);
            }
        }
        return sb.toString();
    }

    /**
     * ��Ԫ��Сƽ����
     * 
     * @param alignment
     * @return
     */
    private String convertAlignToHtml(short alignment) {
        String align = "left";
        switch (alignment) {
        case HSSFCellStyle.ALIGN_LEFT:
            align = "left";
            break;
        case HSSFCellStyle.ALIGN_CENTER:
            align = "center";
            break;
        case HSSFCellStyle.ALIGN_RIGHT:
            align = "right";
            break;
        default:
            break;
        }
        return align;
    }

    /**
     * ��Ԫ��ֱ����
     * 
     * @param verticalAlignment
     * @return
     */
    private String convertVerticalAlignToHtml(short verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
        case HSSFCellStyle.VERTICAL_BOTTOM:
            valign = "bottom";
            break;
        case HSSFCellStyle.VERTICAL_CENTER:
            valign = "center";
            break;
        case HSSFCellStyle.VERTICAL_TOP:
            valign = "top";
            break;
        default:
            break;
        }
        return valign;
    }

    public void makeFile() {
        String sdStateString = android.os.Environment.getExternalStorageState();// ��ȡ�ⲿ�洢״̬
        if (sdStateString.equals(android.os.Environment.MEDIA_MOUNTED)) {// ȷ��sd������,ԭ��֪,ý�尲װ??
            try {
                File sdFile = android.os.Environment
                        .getExternalStorageDirectory();// ��ȡ��չ�豸���ļ�Ŀ¼
                String path = sdFile.getAbsolutePath() + File.separator
                        + "gdemm" + File.separator + "aaa";// �õ�sd��(��չ�豸)�ľ���·��+"/"+xiao
                File dirFile = new File(path);// ��ȡxiao�ļ��е�ַ
                if (!dirFile.exists()) {// ���������
                    dirFile.mkdir();// ����Ŀ¼
                }
                File myFile = new File(path + File.separator + "file.html");// ��ȡmy.html�ĵ�ַ
                if (!myFile.exists()) {// ���������
                    myFile.createNewFile();// �����ļ�
                }
                this.htmlPath = myFile.getAbsolutePath();// ����·��
            } catch (Exception e) {
            }
        }
    }

    /* ������sdcard�ϴ���ͼƬ */
    public void makePictureFile() {
        String sdString = android.os.Environment.getExternalStorageState();// ��ȡ�ⲿ�洢״̬
        if (sdString.equals(android.os.Environment.MEDIA_MOUNTED)) {// ȷ��sd������,ԭ��֪
            try {
                File picFile = android.os.Environment
                        .getExternalStorageDirectory();// ��ȡsd��Ŀ¼
                String picPath = picFile.getAbsolutePath() + File.separator
                        + "gdemm" + File.separator + "aaa";// ����Ŀ¼,�����н���
                File picDirFile = new File(picPath);
                if (!picDirFile.exists()) {
                    picDirFile.mkdir();
                }
                File pictureFile = new File(picPath + File.separator
                        + presentPicture + ".jpg");// ����jpg�ļ�,������html��ͬ
                if (!pictureFile.exists()) {
                    pictureFile.createNewFile();
                }
                this.picturePath = pictureFile.getAbsolutePath();// ��ȡjpg�ļ�����·��
            } catch (Exception e) {
                System.out.println("PictureFile Catch Exception");
            }
        }
    }

    public void writePicture() {
        Picture picture = (Picture) pictures.get(presentPicture);

        byte[] pictureBytes = picture.getContent();

        Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0,
                pictureBytes.length);

        makePictureFile();
        presentPicture++;

        File myPicture = new File(picturePath);

        try {

            FileOutputStream outputPicture = new FileOutputStream(myPicture);

            outputPicture.write(pictureBytes);

            outputPicture.close();
        } catch (Exception e) {
            System.out.println("outputPicture Exception");
        }

        String imageString = "<img src=\"" + picturePath + "\"";
        imageString = imageString + ">";

        try {
            output.write(imageString.getBytes());
        } catch (Exception e) {
            System.out.println("output Exception");
        }
    }

    public int decideSize(int size) {

        if (size >= 1 && size <= 8) {
            return 1;
        }
        if (size >= 9 && size <= 11) {
            return 2;
        }
        if (size >= 12 && size <= 14) {
            return 3;
        }
        if (size >= 15 && size <= 19) {
            return 4;
        }
        if (size >= 20 && size <= 29) {
            return 5;
        }
        if (size >= 30 && size <= 39) {
            return 6;
        }
        if (size >= 40) {
            return 7;
        }
        return 3;
    }

    private String decideColor(int a) {
        int color = a;
        switch (color) {
        case 1:
            return "#000000";
        case 2:
            return "#0000FF";
        case 3:
        case 4:
            return "#00FF00";
        case 5:
        case 6:
            return "#FF0000";
        case 7:
            return "#FFFF00";
        case 8:
            return "#FFFFFF";
        case 9:
            return "#CCCCCC";
        case 10:
        case 11:
            return "#00FF00";
        case 12:
            return "#080808";
        case 13:
        case 14:
            return "#FFFF00";
        case 15:
            return "#CCCCCC";
        case 16:
            return "#080808";
        default:
            return "#000000";
        }
    }

    private void getRange() {
        FileInputStream in = null;
        POIFSFileSystem pfs = null;

        try {
            in = new FileInputStream(nameStr);
            pfs = new POIFSFileSystem(in);
            hwpf = new HWPFDocument(pfs);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        range = hwpf.getRange();

        pictures = hwpf.getPicturesTable().getAllPictures();

        tableIterator = new TableIterator(range);

    }

    public void writeDOCXPicture(byte[] pictureBytes) {
        Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0,
                pictureBytes.length);
        makePictureFile();
        this.presentPicture++;
        File myPicture = new File(this.picturePath);
        try {
            FileOutputStream outputPicture = new FileOutputStream(myPicture);
            outputPicture.write(pictureBytes);
            outputPicture.close();
        } catch (Exception e) {
            System.out.println("outputPicture Exception");
        }
        String imageString = "<img src=\"" + this.picturePath + "\"";

        imageString = imageString + ">";
        try {
            this.output.write(imageString.getBytes());
        } catch (Exception e) {
            System.out.println("output Exception");
        }
    }

    public void writeParagraphContent(Paragraph paragraph) {
        Paragraph p = paragraph;
        int pnumCharacterRuns = p.numCharacterRuns();

        for (int j = 0; j < pnumCharacterRuns; j++) {

            CharacterRun run = p.getCharacterRun(j);

            if (run.getPicOffset() == 0 || run.getPicOffset() >= 1000) {
                if (presentPicture < pictures.size()) {
                    writePicture();
                }
            } else {
                try {
                    String text = run.text();
                    if (text.length() >= 2 && pnumCharacterRuns < 2) {
                        output.write(text.getBytes());
                    } else {
                        int size = run.getFontSize();
                        int color = run.getColor();
                        String fontSizeBegin = "<font size=\""
                                + decideSize(size) + "\">";
                        String fontColorBegin = "<font color=\""
                                + decideColor(color) + "\">";
                        String fontEnd = "</font>";
                        String boldBegin = "<b>";
                        String boldEnd = "</b>";
                        String islaBegin = "<i>";
                        String islaEnd = "</i>";

                        output.write(fontSizeBegin.getBytes());
                        output.write(fontColorBegin.getBytes());

                        if (run.isBold()) {
                            output.write(boldBegin.getBytes());
                        }
                        if (run.isItalic()) {
                            output.write(islaBegin.getBytes());
                        }

                        output.write(text.getBytes());

                        if (run.isBold()) {
                            output.write(boldEnd.getBytes());
                        }
                        if (run.isItalic()) {
                            output.write(islaEnd.getBytes());
                        }
                        output.write(fontEnd.getBytes());
                        output.write(fontEnd.getBytes());
                    }
                } catch (Exception e) {
                    System.out.println("Write File Exception");
                }
            }
        }
    }

    private void readXLSaa() throws Exception {
        List<PictureData> picInfos = new ArrayList<PictureData>();
        myFile = new File(htmlPath);
        output = new FileOutputStream(myFile);
        lsb.append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>");
        lsb.append("<head><meta http-equiv=Content-Type content='text/html; charset=utf-8'><meta name=ProgId content=Excel.Sheet>");
        HSSFSheet sheet = null;
        String excelFileName = nameStr;
        try {
            FileInputStream fis = new FileInputStream(excelFileName);
            HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(fis); // ������Excel
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                if (workbook.getSheetAt(sheetIndex) != null) {
                    sheet = workbook.getSheetAt(sheetIndex);// ��ò�Ϊ�յ����sheet
                    if (picInfos.size() > 0 && picInfos != null) {
                        picInfos.clear();
                    }
                    if (sheet.getDrawingPatriarch() != null) {
                        List<HSSFShape> shapes = sheet.getDrawingPatriarch()
                                .getChildren();
                        for (HSSFShape shape : shapes) {
                            HSSFClientAnchor anchor = (HSSFClientAnchor) shape
                                    .getAnchor();
                            if (shape instanceof HSSFPicture) {
                                HSSFPicture pic = (HSSFPicture) shape;
                                pic.getPictureData();
                                PictureData info = new PictureData(workbook, sheet, null, null);

                                int row = anchor.getRow1();
                                int col = anchor.getCol1();
                                info.setColNum(col);
                                info.setRowNum(row);
                                info.setSheetName(sheet.getSheetName());
                                HSSFPictureData picData = pic.getPictureData();
                                try {
                                    info = savePic(info, sheet.getSheetName(),
                                            picData);
                                } catch (Exception e) {
                                    e.printStackTrace();
                                }
                                picInfos.add(info);
                            }
                        }

                    }
                    int firstRowNum = sheet.getFirstRowNum(); // ��һ��
                    int lastRowNum = sheet.getLastRowNum(); // ���һ��
                    // ����Table
                    lsb.append("<table width=\"100%\" style=\"border:1px solid #000;border-width:1px 0 0 1px;margin:2px 0 2px 0;border-collapse:collapse;\">");
                    for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
                        if (sheet.getRow(rowNum) != null) {// ����в�Ϊ�գ�
                            HSSFRow row = sheet.getRow(rowNum);
                            short firstCellNum = row.getFirstCellNum(); // ���еĵ�һ����Ԫ��
                            short lastCellNum = row.getLastCellNum(); // ���е����һ����Ԫ��
                            int height = (int) (row.getHeight() / 15.625); // �еĸ߶�
                            lsb.append("<tr height=\""
                                    + height
                                    + "\" style=\"border:1px solid #000;border-width:0 1px 1px 0;margin:2px 0 2px 0;\">");
                            for (short cellNum = firstCellNum; cellNum <= lastCellNum; cellNum++) { // ѭ�����е�ÿһ����Ԫ��
                                HSSFCell cell = row.getCell(cellNum);
                                StringBuffer tdStyle = new StringBuffer(
                                        "<td style=\"border:1px solid #000; border-width:0 1px 1px 0;margin:2px 0 2px 0; ");
                                if (null != cell) {
                                    HSSFCellStyle cellStyle = cell
                                            .getCellStyle();

                                    short boldWeight = cellStyle.getFont(
                                            workbook).getBoldweight(); // �����ϸ
                                    short fontHeight = (short) (cellStyle
                                            .getFont(workbook).getFontHeight() / 2); // �����С
                                    tdStyle.append(" font-weight:" + boldWeight
                                            + "; ");
                                    tdStyle.append(" font-size: " + fontHeight
                                            + "%;");
                                    lsb.append(tdStyle + "\"");

                                }
                                boolean flag = false;
                                if (picInfos.size() > 0 && picInfos != null) {
                                    for (PictureData picInfo : picInfos) {
                                        // �ҵ�ͼƬ��Ӧ�ĵ�Ԫ��
                                        if (picInfo.getSheetName().equals(
                                                sheet.getSheetName())
                                                && picInfo.getRowNum() == rowNum
                                                && picInfo.getColNum() == cellNum) {
                                            flag = true;
                                            String imagePath = "<img src=\""
                                                    + picInfo.getPicPath()
                                                    + "\"" + "/>";
                                            lsb.append("<td>");
                                            lsb.append(">" + imagePath
                                                    + "</td>");
                                        }
                                    }
                                }
                                if (null != cell && !flag) {
                                    lsb.append(">" + getCellValue(cell)
                                            + "</td>");
                                }
                            }
                            lsb.append("</tr>");
                        }
                    }

                }
            }
            output.write(lsb.toString().getBytes());
            fis.close();
        } catch (FileNotFoundException e) {
            throw new Exception("�ļ� " + excelFileName + " û���ҵ�!");
        } catch (IOException e) {
            throw new Exception("�ļ� " + excelFileName + " �������("
                    + e.getMessage() + ")!");
        }

    }

    private PictureData savePic(PictureData info, String sheetName,
            HSSFPictureData picData) {
        
        return null;
    }

}
