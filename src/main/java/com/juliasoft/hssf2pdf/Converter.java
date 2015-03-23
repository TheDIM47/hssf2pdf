package com.juliasoft.hssf2pdf;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.awt.*;
import java.io.*;

public class Converter {

    public void convert() throws IOException, DocumentException, InvalidFormatException {
        final File fontFile = new File(System.getProperty("user.dir"), "Arial.ttf");
        final File[] samples = new File(System.getProperty("user.dir"), "/samples/").listFiles();
        for(final File src : samples) {
            final File target = new File(System.getProperty("user.dir"), src.getName() + ".pdf");
            final Document document = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
            try (final OutputStream os = new FileOutputStream(target)) {
                final PdfWriter writer = PdfWriter.getInstance(document, os);
                try {
                    document.open();
                    try (final InputStream is = new FileInputStream(src)) {
                        new HssfRenderer(document, new FontFactory(fontFile), is).withAdjustHeight(true).withAdjustWidth(true).withBorders(false).render();
                    } finally {
                        document.close();
                    }
                } finally {
                    writer.flush();
                    os.flush();
                }
            }
            Desktop.getDesktop().open(target);
        }
    }

    public static void main(String[] args) throws IOException, DocumentException, InvalidFormatException {
        new Converter().convert();
    }
}
