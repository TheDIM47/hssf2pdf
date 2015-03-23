package com.juliasoft.hssf2pdf;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPRow;
import com.itextpdf.text.pdf.PdfPTable;
import com.juliasoft.utils.FontInfo;
import com.juliasoft.utils.HSSF2PDFUtils;
import com.juliasoft.utils.ImageInfo;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;

public class HssfRenderer {
    public HssfRenderer(Document doc, FontFactory fontFactory, InputStream is) {
        this.doc = doc;
        this.is = is;
        this.fontFactory = fontFactory;
    }

    public HssfRenderer withBorders(boolean withBorders) {
        this.withBorders = withBorders;
        return this;
    }

    public HssfRenderer withAdjustWidth() {
        return withAdjustWidth(true);
    }

    public HssfRenderer withAdjustHeight() {
        return withAdjustHeight(true);
    }

    public HssfRenderer withAdjustWidth(boolean withAdjustWidth) {
        this.adjustWidth = withAdjustWidth;
        return this;
    }

    public HssfRenderer withAdjustHeight(boolean withAdjustHeight) {
        this.adjustHeight = withAdjustHeight;
        return this;
    }

    public HssfRenderer withTextPadding(float left, float right, float top, float bottom) {
        this.paddingLeft = left;
        this.paddingRight = right;
        this.paddingTop = top;
        this.paddingBottom = bottom;
        return this;
    }

    public Document render() throws IOException, InvalidFormatException, DocumentException {
        final HSSFWorkbook wb = new HSSFWorkbook(is);
        final HSSFSheet sheet = wb.getSheetAt(0);

        if (sheet.getLastRowNum() == 0) {
            return doc;
        }
        int minCell = 0, maxCell = 0;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            final HSSFRow row = sheet.getRow(i);
            if (row != null && row.getFirstCellNum() != -1) {
                minCell = Math.min(minCell, row.getFirstCellNum());
                maxCell = Math.max(maxCell, row.getLastCellNum());
            }
        }
        if (minCell == maxCell && maxCell == 0) {
            return doc;
        }

        final HSSFCellStyle ZERO_STYLE = wb.createCellStyle();
        ZERO_STYLE.setBorderLeft((short) 0);
        ZERO_STYLE.setBorderRight((short) 0);
        ZERO_STYLE.setBorderTop((short) 0);
        ZERO_STYLE.setBorderBottom((short) 0);

        final Iterable<CellRangeAddress> merges = HSSF2PDFUtils.getMergedRegions(sheet);

        final Rectangle pageSize = doc.getPageSize();

        final float[] widths;
        if (adjustWidth) {
            final float margins = 1 + doc.leftMargin() + doc.rightMargin();
            widths = HSSF2PDFUtils.adjust(HSSF2PDFUtils.getColumnWidths(sheet, minCell, maxCell), pageSize.getWidth() - margins);
        } else {
            widths = HSSF2PDFUtils.getColumnWidths(sheet, minCell, maxCell);
        }

        final float[] heights;
        if (adjustHeight) {
            final float margins = 1 + doc.topMargin() + doc.bottomMargin();
            heights = HSSF2PDFUtils.adjust(HSSF2PDFUtils.getRowHeights(sheet), pageSize.getHeight() - margins);
        } else {
            heights = HSSF2PDFUtils.getRowHeights(sheet);
        }

        final PdfPTable table = new PdfPTable(maxCell - minCell);
        table.setWidthPercentage(100); // TODO: Fixme!
        table.setWidths(widths);

        table.getDefaultCell().setFixedHeight(10f);

        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            final HSSFRow row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int colIndex = minCell; colIndex < row.getLastCellNum(); colIndex++) {
                    final PdfPCell pdfCell = HSSF2PDFUtils.getCell(merges, rowIndex, colIndex);
                    if (pdfCell == null) {
                        continue;
                    }
                    // Excel cell
                    final HSSFCell xlCell = row.getCell(colIndex);
                    if (xlCell != null) {
                        final Object value = HSSF2PDFUtils.getCellValue(xlCell);
                        final String strValue = String.valueOf(value);

                        cloneStyle(pdfCell, xlCell.getCellStyle(), withBorders);

                        final com.itextpdf.text.Font font = getFont(wb, xlCell.getCellStyle());
                        pdfCell.setPhrase(new Phrase(strValue, font));

                        // calc cell height
                        float cellHeight = 0;
                        for (int i = rowIndex; i < rowIndex + pdfCell.getRowspan(); i++) {
                            cellHeight += heights[i];
                        }
                        pdfCell.setFixedHeight(cellHeight);
                    } else {
                        cloneStyle(pdfCell, ZERO_STYLE, withBorders);
                    }
                    table.addCell(pdfCell);
                }
                table.completeRow();
            } else {
                table.completeRow();
            }
        }

        // process images
        final java.util.List<ImageInfo> images = HSSF2PDFUtils.getImages(sheet);
        for (ImageInfo info : images) {
            final PdfPRow row = table.getRow(info.getRow1());
            final PdfPCell[] cells = row.getCells();
            int counter = 0;
            for (PdfPCell cell : cells) {
                if (cell != null) {
                    final int colspan = cell.getColspan();
                    if ((info.getCol1() >= counter) && (info.getCol1() < (counter + colspan))) {
                        final Image img = Image.getInstance(info.getImageData());
                        cell.setImage(img);
                        break;
                    }
                    counter += colspan;
                }
            }
        }

        doc.add(table);
        return doc;
    }

    private void cloneStyle(PdfPCell pdfCell, HSSFCellStyle xlStyle, boolean withBorders) throws IOException, DocumentException {
        if (pdfCell == null) {
            throw new IllegalArgumentException("PDF Cell object must not be null");
        }
        if (xlStyle == null) {
            throw new IllegalArgumentException("Excel Cell Style object must not be null");
        }
        short rotation = xlStyle.getRotation();
        if (rotation != 0) {
            if (rotation > 90) {
                rotation = (short) (450 - rotation); // (360 + (90 - rotation))
            }
            pdfCell.setRotation(rotation);
        }
        pdfCell.setHorizontalAlignment(HSSF2PDFUtils.cvtHAlign(xlStyle.getAlignment()));
        pdfCell.setVerticalAlignment(HSSF2PDFUtils.cvtVAlign(xlStyle.getVerticalAlignment()));

        final HSSFColor bColor = HSSFColor.getIndexHash().get((int) xlStyle.getFillBackgroundColor());
        final BaseColor backColor = HSSF2PDFUtils.rgbToBaseColor(bColor == null ? HSSFColor.WHITE.triplet : bColor.getTriplet());

        final HSSFColor fColor = HSSFColor.getIndexHash().get((int) xlStyle.getFillForegroundColor());
        final BaseColor foreColor = HSSF2PDFUtils.rgbToBaseColor(fColor == null ? HSSFColor.WHITE.triplet : fColor.getTriplet());

        if (bColor != null && fColor != null) {
            final BaseColor summaryColor = HSSF2PDFUtils.mixColors(backColor, foreColor);
            pdfCell.setBackgroundColor(summaryColor);
        } else if (bColor != null) {
            pdfCell.setBackgroundColor(backColor);
        } else if (fColor != null) {
            pdfCell.setBackgroundColor(foreColor);
        }

        pdfCell.setNoWrap(!xlStyle.getWrapText());

        pdfCell.setPaddingLeft(paddingLeft);
        pdfCell.setPaddingRight(paddingRight);
        pdfCell.setPaddingTop(paddingTop);
        pdfCell.setPaddingBottom(paddingBottom);

        pdfCell.setUseAscender(true);
        pdfCell.setUseDescender(true);

        if (withBorders) {
            pdfCell.setBorderWidthLeft(HSSF2PDFUtils.cvtBorder(xlStyle.getBorderLeft()));
            pdfCell.setBorderWidthRight(HSSF2PDFUtils.cvtBorder(xlStyle.getBorderRight()));
            pdfCell.setBorderWidthTop(HSSF2PDFUtils.cvtBorder(xlStyle.getBorderTop()));
            pdfCell.setBorderWidthBottom(HSSF2PDFUtils.cvtBorder(xlStyle.getBorderBottom()));

            pdfCell.setBorderColorLeft(HSSF2PDFUtils.hssfToBaseColor(HSSFColor.getIndexHash().get((int) xlStyle.getLeftBorderColor())));
            pdfCell.setBorderColorRight(HSSF2PDFUtils.hssfToBaseColor(HSSFColor.getIndexHash().get((int) xlStyle.getRightBorderColor())));
            pdfCell.setBorderColorTop(HSSF2PDFUtils.hssfToBaseColor(HSSFColor.getIndexHash().get((int) xlStyle.getTopBorderColor())));
            pdfCell.setBorderColorBottom(HSSF2PDFUtils.hssfToBaseColor(HSSFColor.getIndexHash().get((int) xlStyle.getBottomBorderColor())));
        } else {
            pdfCell.setBorderWidthLeft(0);
            pdfCell.setBorderWidthRight(0);
            pdfCell.setBorderWidthTop(0);
            pdfCell.setBorderWidthBottom(0);
        }
    }

    private com.itextpdf.text.Font getFont(HSSFWorkbook wb, HSSFCellStyle cellStyle) throws IOException, DocumentException {
        if (wb == null) {
            throw new IllegalArgumentException("PDF Cell object must not be null");
        }
        if (cellStyle == null) {
            throw new IllegalArgumentException("Cell Style object must not be null");
        }
        final FontInfo fontInfo = HSSF2PDFUtils.getFontInfo(wb, cellStyle);
        return fontFactory.getFont(fontInfo);
    }


    private boolean withBorders = false;
    private boolean adjustWidth = true;
    private boolean adjustHeight = true;
    private float paddingLeft = 1f;
    private float paddingRight = 1f;
    private float paddingTop = 1f;
    private float paddingBottom = 1f;

    private final FontFactory fontFactory;
    private final Document doc;
    private final InputStream is;

    private static final Logger log = LoggerFactory.getLogger(HssfRenderer.class);
}
