package com.zjf.util;

import java.io.File;

import org.apache.poi.ddf.EscherClientAnchorRecord;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class PictureData {
    private final HSSFWorkbook workbook;
    private final HSSFSheet sheet;
    private final HSSFPictureData pictureData;
    private final EscherClientAnchorRecord clientAnchor;
    int ColNum;
    int RowNum;
    String SheetName;
    String picPath;
    
    public String getSheetName() {
        return SheetName;
    }
    
    public String  getPicPath(){
        File sdFile = android.os.Environment
                .getExternalStorageDirectory();// 获取扩展设备的文件目录
        String path = sdFile.getAbsolutePath() + File.separator
                + "gdemm" + File.separator + "aaa";
        return path;
        
    }


    public void setSheetName(String sheetName) {
        SheetName = sheetName;
    }


    public int getRowNum() {
        return RowNum;
    }


    public void setRowNum(int rowNum) {
        RowNum = rowNum;
    }


    public int getColNum() {
        return ColNum;
    }


    public void setColNum(int colNum) {
        ColNum = colNum;
    }


    public PictureData(HSSFWorkbook workbook, HSSFSheet sheet, HSSFPictureData pictureData, EscherClientAnchorRecord clientAnchor) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.pictureData = pictureData;
        this.clientAnchor = clientAnchor;
    }
    

    public HSSFWorkbook getWorkbook() {
        return workbook;
    }
    
    public HSSFSheet getSheet() {
        return sheet;
    }
    
    public EscherClientAnchorRecord getClientAnchor() {
        return clientAnchor;
    }
    
    public HSSFPictureData getPictureData() {
        return pictureData;
    }

    public byte[] getData() {
        return pictureData.getData();
    }

    public String suggestFileExtension() {
        return pictureData.suggestFileExtension();
    }
    
    /**
     * 推测图片中心所覆盖的单元格，这个值不一定准确，但通常有效
     * 
     * @return the row0
     */
    public short getRow0() {
        int row1 = getRow1();
        int row2 = getRow2();
        if (row1 == row2) {
            return (short) row1;
        }
        
        int heights[] = new int[row2-row1+1];
        for (int i = 0; i < heights.length; i++) {
            heights[i] = getRowHeight(row1 + i);
        }
        
        // HSSFClientAnchor 中 dx 只能在 0-1023 之间,dy 只能在 0-255 之间
        // 表示相对位置的比率，不是绝对值
        int dy1 = getDy1() * heights[0] / 255;
        int dy2 = getDy2() * heights[heights.length-1] / 255;
        return (short) (getCenter(heights, dy1, dy2) + row1);
    }
    
    
    private short getRowHeight(int rowIndex) {
        HSSFRow row = sheet.getRow(rowIndex);
        short h = row == null? sheet.getDefaultRowHeight() : row.getHeight();
        return h;
    }
    
    /**
     * 推测图片中心所覆盖的单元格，这个值不一定准确，但通常有效
     * 
     * @return the col0
     */
    public short getCol0() {
        short col1 = getCol1();
        short col2 = getCol2();
        
        if (col1 == col2) {
            return col1;
        }
        
        int widths[] = new int[col2-col1+1];
        for (int i = 0; i < widths.length; i++) {
            widths[i] = sheet.getColumnWidth(col1 + i);
        }
        
        // HSSFClientAnchor 中 dx 只能在 0-1023 之间,dy 只能在 0-255 之间
        // 表示相对位置的比率，不是绝对值
        int dx1 = getDx1() * widths[0] / 1023;
        int dx2 = getDx2() * widths[widths.length-1] / 1023;

        return (short) (getCenter(widths, dx1, dx2) + col1);
    }
    
    /**
     * 给定各线段的长度，以及起点相对于起点段的偏移量，终点相对于终点段的偏移量，
     * 求中心点所在的线段
     * 
     * @param a the a 各线段的长度
     * @param d1 the d1 起点相对于起点段
     * @param d2 the d2 终点相对于终点段的偏移量
     * 
     * @return the center
     */
    protected static int getCenter(int[] a, int d1, int d2) {
        // 线段长度
        int width = a[0] - d1 + d2;
        for (int i = 1; i < a.length-1; i++) {
            width += a[i];
        }

        // 中心点位置
        int c = width / 2 + d1;
        int x = a[0];
        int cno = 0;
        
        while (c > x) {
            x += a[cno];
            cno++;
        }
        
        return cno;
    }

    /**
     * 左上角所在列
     * 
     * @return the col1
     */
    public short getCol1() {
        return clientAnchor.getCol1();
    }

    /**
     * 右下角所在的列
     * 
     * @return the col2
     */
    public short getCol2() {
        return clientAnchor.getCol2();
    }

    /**
     * 左上角的相对偏移量
     * 
     * @return the dx1
     */
    public short getDx1() {
        return clientAnchor.getDx1();
    }

    /**
     * 右下角的相对偏移量
     * 
     * @return the dx2
     */
    public short getDx2() {
        return clientAnchor.getDx2();
    }

    /**
     * 左上角的相对偏移量
     * 
     * @return the dy1
     */
    public short getDy1() {
        return clientAnchor.getDy1();
    }

    /**
     * 右下角的相对偏移量
     * 
     * @return the dy2
     */
    public short getDy2() {
        return clientAnchor.getDy2();
    }

    /**
     * 左上角所在的行
     * 
     * @return the row1
     */
    public short getRow1() {
        return clientAnchor.getRow1();
    }

    /**
     * 右下角所在的行
     * 
     * @return the row2
     */
    public short getRow2() {
        return clientAnchor.getRow2();
    }
    
}
