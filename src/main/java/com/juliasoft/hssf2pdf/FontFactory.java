package com.juliasoft.hssf2pdf;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.BaseFont;
import com.juliasoft.utils.FontInfo;
import com.juliasoft.utils.HSSF2PDFUtils;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class FontFactory {
    public static final String CP_1251 = "cp1251";
    public static final String CP_UTF8 = "utf-8";

    public FontFactory(File fontFile) throws IOException, DocumentException {
        this(BaseFont.createFont(fontFile.getCanonicalPath(), CP_1251, BaseFont.EMBEDDED));
    }

    public FontFactory(BaseFont bf) {
        this.bf = bf;
    }

    public Font getFont(FontInfo info) {
        if (info == null) {
            throw new IllegalArgumentException("Font Info must not be null");
        }
        if (fontMap.containsKey(info)) {
            return fontMap.get(info);
        } else {
//            final BaseFont bf = BaseFont.createFont(new File(System.getProperty("user.dir"), "Arial.ttf").getCanonicalPath(), CP_1251, BaseFont.EMBEDDED);
            final Font font = new Font(bf, info.getSize());
            font.setFamily(Font.FontFamily.HELVETICA.name());
            font.setStyle(info.getStyle());
            final BaseColor color = HSSF2PDFUtils.awtToBaseColor(info.getColor());
            font.setColor(color);
            fontMap.put(info, font);
            return font;
        }
    }

    private final BaseFont bf;
    private Map<FontInfo, Font> fontMap = new HashMap<>();
}
