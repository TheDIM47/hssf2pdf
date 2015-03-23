package com.juliasoft.utils;

import java.awt.*;

public class FontInfo {

    public FontInfo(String name, int size, int style, Color color) {
        this.name = name;
        this.size = size;
        this.style = style;
        this.color = color;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getSize() {
        return size;
    }

    public void setSize(int size) {
        this.size = size;
    }

    public int getStyle() {
        return style;
    }

    public void setStyle(int style) {
        this.style = style;
    }

    public Color getColor() {
        return color;
    }

    public void setColor(Color color) {
        this.color = color;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof FontInfo)) return false;

        FontInfo fontInfo = (FontInfo) o;

        if (size != fontInfo.size) return false;
        if (style != fontInfo.style) return false;
        if (!color.equals(fontInfo.color)) return false;
        if (!name.equals(fontInfo.name)) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = name.hashCode();
        result = 31 * result + size;
        result = 31 * result + style;
        result = 31 * result + color.hashCode();
        return result;
    }

    private String name;
    private int size;
    private int style;
    private Color color;
}
