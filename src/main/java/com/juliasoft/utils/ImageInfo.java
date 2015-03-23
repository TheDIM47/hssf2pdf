package com.juliasoft.utils;

public class ImageInfo {

    public int getCol1() {
        return col1;
    }

    public void setCol1(int col1) {
        this.col1 = col1;
    }

    public int getRow1() {
        return row1;
    }

    public void setRow1(int row1) {
        this.row1 = row1;
    }

    public int getCol2() {
        return col2;
    }

    public void setCol2(int col2) {
        this.col2 = col2;
    }

    public int getRow2() {
        return row2;
    }

    public void setRow2(int row2) {
        this.row2 = row2;
    }

    public int getImageType() {
        return imageType;
    }

    public void setImageFormat(int imageType) {
        this.imageType = imageType;
    }

    public String getMimeType() {
        return mimeType;
    }

    public void setMimeType(String mimeType) {
        this.mimeType = mimeType;
    }

    public byte[] getImageData() {
        return imageData;
    }

    public void setImageData(byte[] imageData) {
        this.imageData = imageData;
    }

    public boolean isHFlip() {
        return HFlip;
    }

    public void setHFlip(boolean HFlip) {
        this.HFlip = HFlip;
    }

    public boolean isVFlip() {
        return VFlip;
    }

    public void setVFlip(boolean VFlip) {
        this.VFlip = VFlip;
    }

    public int getDx1() {
        return dx1;
    }

    public void setDx1(int dx1) {
        this.dx1 = dx1;
    }

    public int getDy1() {
        return dy1;
    }

    public void setDy1(int dy1) {
        this.dy1 = dy1;
    }

    public int getDx2() {
        return dx2;
    }

    public void setDx2(int dx2) {
        this.dx2 = dx2;
    }

    public int getDy2() {
        return dy2;
    }

    public void setDy2(int dy2) {
        this.dy2 = dy2;
    }

    private int col1, row1, col2, row2;
//    private int x1, y1, x2, y2;
    private int dx1, dy1, dx2, dy2;
    private int imageType;
    private String mimeType;
    private byte[] imageData;
    private boolean HFlip;
    private boolean VFlip;
}
