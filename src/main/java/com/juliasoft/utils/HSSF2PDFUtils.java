package com.juliasoft.utils;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Element;
import com.itextpdf.text.pdf.PdfPCell;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.util.*;
import java.util.List;

public class HSSF2PDFUtils {
    public static BaseColor awtToBaseColor(Color color) {
        return new BaseColor(color.getRed(), color.getGreen(), color.getBlue(), color.getAlpha());
    }

    public static BaseColor rgbToBaseColor(short[] rgb) {
        return new BaseColor(rgb[0], rgb[1], rgb[2]);
    }

    public static BaseColor hssfToBaseColor(HSSFColor color) {
        return rgbToBaseColor(color == null ? HSSFColor.WHITE.triplet : color.getTriplet());
    }

    public static <T> boolean sameValues(T[] arr) {
        if (arr != null && arr.length > 0) {
            final T v = arr[0];
            for (int i = 1; i < arr.length; i++) {
                if (!v.equals(arr[i])) {
                    return false;
                }
            }
        }
        return true;
    }

    public static boolean sameValues(short[] arr) {
        if (arr != null && arr.length > 0) {
            final short v = arr[0];
            for (int i = 1; i < arr.length; i++) {
                if (v != arr[i]) {
                    return false;
                }
            }
        }
        return true;
    }

    public static float[] adjust(float[] arr, float factor) {
        float total = sum(arr);
        if (arr != null && arr.length > 0 && total > 0) {
            final float[] res = new float[arr.length];
            for (int i = 0; i < arr.length; i++) {
                res[i] = factor * arr[i] / total;
            }
            return res;
        } else {
            return arr;
        }
    }

    public static float sum(float[] arr) {
        float s = 0;
        if (arr != null && arr.length > 0) {
            for (int i = 0; i < arr.length; i++) {
                s += arr[i];
            }
        }
        return s;
    }

    public static BaseColor mixColors(BaseColor bg, BaseColor fg) {
        float fgA = fg.getAlpha() / 255f;
        float bgA = bg.getAlpha() / 255f;
        final float rA = 1f - (1f - fgA) * (1f - bgA);
        final float rR = fg.getRed() * fgA / rA + bg.getRed() * bgA * (1f - fgA) / rA;
        final float rG = fg.getGreen() * fgA / rA + bg.getGreen() * bgA * (1f - fgA) / rA;
        final float rB = fg.getBlue() * fgA / rA + bg.getBlue() * bgA * (1f - fgA) / rA;
        return new BaseColor(rR / 255f, rG / 255f, rB / 255f, rA);
    }

    public static int cvtHAlign(short hAlign) {
        switch (hAlign) {
            case HSSFCellStyle.ALIGN_LEFT:
                return Element.ALIGN_LEFT;
            case HSSFCellStyle.ALIGN_CENTER:
                return Element.ALIGN_CENTER;
            case HSSFCellStyle.ALIGN_RIGHT:
                return Element.ALIGN_RIGHT;
            case HSSFCellStyle.ALIGN_FILL:
                return Element.ALIGN_JUSTIFIED_ALL;
            case HSSFCellStyle.ALIGN_JUSTIFY:
                return Element.ALIGN_JUSTIFIED;
            case HSSFCellStyle.ALIGN_CENTER_SELECTION:
                return Element.ALIGN_MIDDLE;
            default:
                return Element.ALIGN_UNDEFINED;
        }
    }

    public static int cvtVAlign(short vAlign) {
        switch (vAlign) {
            case HSSFCellStyle.VERTICAL_TOP:
                return Element.ALIGN_TOP;
            case HSSFCellStyle.VERTICAL_CENTER:
                return Element.ALIGN_MIDDLE;
            case HSSFCellStyle.VERTICAL_BOTTOM:
                return Element.ALIGN_BOTTOM;
            case HSSFCellStyle.VERTICAL_JUSTIFY:
                return Element.ALIGN_JUSTIFIED;
            default:
                return Element.ALIGN_MIDDLE;
        }
    }

    public static float cvtBorder(short border) {
        final float factor = 0.5f;
        switch (border) {
            case HSSFCellStyle.BORDER_NONE:
                return 0;
            case HSSFCellStyle.BORDER_MEDIUM:
                return 2f * factor;
            case HSSFCellStyle.BORDER_THICK:
                return 3f * factor;
            case HSSFCellStyle.BORDER_DOUBLE:
                return 4f * factor;
            case HSSFCellStyle.BORDER_HAIR:
            case HSSFCellStyle.BORDER_DOTTED:
            case HSSFCellStyle.BORDER_DASHED:
            case HSSFCellStyle.BORDER_MEDIUM_DASHED:
            case HSSFCellStyle.BORDER_DASH_DOT:
            case HSSFCellStyle.BORDER_MEDIUM_DASH_DOT:
            case HSSFCellStyle.BORDER_DASH_DOT_DOT:
            case HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
            case HSSFCellStyle.BORDER_SLANTED_DASH_DOT:
                return factor;
            default:
                return factor; // BORDER_THIN
        }
    }

    /**
     * Create and return cell or null if cell inside merged region (but not at start of merged region)
     * Set rowspan and colspan if cell is start of merged region
     *
     * @param merges set of merged regions
     * @param row    cell row
     * @param col    cell column
     * @return pdf cell or null if cell not need to be created
     */
    public static PdfPCell getCell(Iterable<CellRangeAddress> merges, int row, int col) {
        if (merges == null) {
            throw new IllegalArgumentException("Cell Range Addresses must not be null");
        }
        final PdfPCell cell = new PdfPCell();
        for (CellRangeAddress range : merges) {
            if (range.isInRange(row, col)) {
                if (range.getFirstRow() == row && range.getFirstColumn() == col) {
                    final int width = 1 + range.getLastColumn() - range.getFirstColumn();
                    final int height = 1 + range.getLastRow() - range.getFirstRow();
                    if (width > 1) {
                        cell.setColspan(width);
                    }
                    if (height > 1) {
                        cell.setRowspan(height);
                    }
                    break;
                } else {
                    return null;
                }
            }
        }
        return cell;
    }

    /**
     * Get Sheet columns width
     *
     * @param sheet   current sheet
     * @param minCell minimal cell number
     * @param maxCell maximal cell number
     * @return column widths (in units of 1/256th of a character width)
     */
    public static float[] getColumnWidths(HSSFSheet sheet, int minCell, int maxCell) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet must not be null");
        }
        final float widths[] = new float[maxCell - minCell];
        for (int i = minCell; i < maxCell; i++) {
            widths[i] = sheet.getColumnWidth(i);
        }
        return widths;
    }

    public static float[] getRowHeights(HSSFSheet sheet) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet must not be null");
        }
        final float heights[] = new float[sheet.getLastRowNum() + 1];
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            final Row row = sheet.getRow(i);
            if (row != null) {
                // n / 100.0 / 25.4 * 72.0
                heights[i] = row.getHeightInPoints();
            }
        }
        return heights;
    }

    /**
     * Get merged regions from sheet
     *
     * @param sheet current sheet
     * @return set of CellRangeAddress of merged regions
     */
    public static Iterable<CellRangeAddress> getMergedRegions(HSSFSheet sheet) {
        final Set<CellRangeAddress> merges = new HashSet<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            final CellRangeAddress range = sheet.getMergedRegion(i);
            merges.add(range);
        }
        return merges;
    }

    /**
     * Extract FontInfo from cell style
     *
     * @param wb
     * @param cellStyle
     * @return
     */
    public static FontInfo getFontInfo(HSSFWorkbook wb, HSSFCellStyle cellStyle) {
        final short fontIndex = cellStyle.getFontIndex();
        final org.apache.poi.ss.usermodel.Font f = wb.getFontAt(fontIndex);

        final String fontName = f.getFontName();
        final int fontSize = f.getFontHeightInPoints();
        final HSSFColor hssfColor = HSSFColor.getIndexHash().get((int) f.getColor());
        final short[] rgb = (hssfColor == null) ? HSSFColor.BLACK.triplet : hssfColor.getTriplet();
        final java.awt.Color awtColor = new java.awt.Color(rgb[0], rgb[1], rgb[2]);

        int fontStyle = 0;
        if (f.getBoldweight() > 400) {
            fontStyle |= com.itextpdf.text.Font.BOLD;
        }
        if (f.getItalic()) {
            fontStyle |= com.itextpdf.text.Font.ITALIC;
        }
        if (f.getUnderline() > 0) {
            fontStyle |= com.itextpdf.text.Font.UNDERLINE;
        }
        if (f.getStrikeout()) {
            fontStyle |= com.itextpdf.text.Font.STRIKETHRU;
        }
        return new FontInfo(fontName, fontSize, fontStyle, awtColor);
    }

    public static List<ImageInfo> getImages(HSSFSheet sheet) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet must not be null");
        }
        final List<ImageInfo> images = new ArrayList<>();
        final HSSFPatriarch patriarch = sheet.getDrawingPatriarch();
        if (patriarch != null) {
            final java.util.List<HSSFShape> shapes = patriarch.getChildren();
            log.trace("X1:{}, Y1:{}, X2:{}, Y2:{}, Shapes:{}", patriarch.getX1(), patriarch.getY1(), patriarch.getX2(), patriarch.getY2(), shapes);
            if (shapes != null) {
                log.trace("Shapes size: {}", shapes.size());
                for (HSSFShape shape : shapes) {
                    final ImageInfo info = new ImageInfo();

                    final HSSFAnchor ar = shape.getAnchor();
                    log.trace("Shape: {}, Anchor: {}, childs: {}", shape, ar, shape.countOfAllChildren());
                    log.trace("Anchor: DX1:{}, DY1:{}, DX2:{}, DY2:{}", ar.getDx1(), ar.getDy1(), ar.getDx2(), ar.getDy2());

                    info.setDx1(ar.getDx1());
                    info.setDy1(ar.getDy1());
                    info.setDx2(ar.getDx2());
                    info.setDy2(ar.getDy2());

                    if (ar instanceof HSSFClientAnchor) {
                        final HSSFClientAnchor car = (HSSFClientAnchor) ar;
                        log.trace("Client Anchor CR: [{},{}:{},{}]", car.getCol1(), car.getRow1(), car.getCol2(), car.getRow2());

                        info.setHFlip(car.isHorizontallyFlipped());
                        info.setVFlip(car.isVerticallyFlipped());

                        info.setCol1(car.getCol1());
                        info.setRow1(car.getRow1());
                        info.setCol2(car.getCol2());
                        info.setRow2(car.getRow2());
                    }
                    if (shape instanceof HSSFPicture) {
                        final HSSFPicture picture = (HSSFPicture) shape;
                        final HSSFPictureData pic = picture.getPictureData();
                        log.trace("Picture data: Fmt:{}, Mime: {}, Data len: {}", pic.getFormat(), pic.getMimeType(), pic.getData().length);

                        info.setImageFormat(pic.getFormat());
                        info.setImageData(pic.getData());
                        info.setMimeType(pic.getMimeType());
//                        final Image img = Image.getInstance(pic.getData());
//                        log.trace("Got Image: {}", img);
                    }
                    images.add(info);
                }
            }
        }
        return images;
    }

    public static String getFormatName(int format) {
        switch (format) {
            case HSSFWorkbook.PICTURE_TYPE_EMF: // Windows Enhanced Metafile
                return "EMF";
            case HSSFWorkbook.PICTURE_TYPE_WMF: // Windows Metafile
                return "WMF";
            case HSSFWorkbook.PICTURE_TYPE_PICT: // Macintosh PICT
                return "PICT";
            case HSSFWorkbook.PICTURE_TYPE_JPEG: // JFIF
                return "JFIF";
            case HSSFWorkbook.PICTURE_TYPE_PNG: // PNG
                return "PNG";
            case HSSFWorkbook.PICTURE_TYPE_DIB: // Windows DIB
                return "DIB";
            default:
                throw new IllegalArgumentException("Unknown picture type: " + format);
        }
    }

    public static void setCellValue(Sheet sheet, int row, int col, String value) {
        if (value != null) {
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellValue(value);
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
        }
    }

    public static void setCellValue(Sheet sheet, int row, int col, Date value) {
        if (value != null) {
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellValue(value);
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
        }
    }

    public static void setCellValue(Sheet sheet, int row, int col, Boolean value) {
        if (value != null) {
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellValue(value);
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_BOOLEAN);
        }
    }

    public static void setCellValue(Sheet sheet, int row, int col, Number value) {
        if (value != null) {
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellValue(value.doubleValue());
        } else {
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellValue(0);
        }
        sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_NUMERIC);
    }

    public static void setCellValue(Sheet sheet, int row, int col, BigDecimal value) {
        if (value != null) {
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellValue(value.doubleValue());
        } else {
            sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellValue(0);
        }
        sheet.getRow(row).getCell(col, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_NUMERIC);
    }

    public static Object getCellValue(final Cell cell) {
        if (cell == null) {
            return null;
        } else {
            return getCellValue(cell, cell.getCellType());
        }
    }

    /**
     * Вернет значение данных в ячейке. Если ячейка содержит формулу - вернет кэшированные данные: результат вычисления формулы.
     *
     * @param cell      тестируемая ячейка
     * @param valueType тип значения
     * @return значение данных в ячейке
     * @throws Exception
     */
    public static Object getCellValue(final Cell cell, final int valueType) {
        switch (valueType) {
            case Cell.CELL_TYPE_FORMULA:
                return getCellValue(cell, cell.getCachedFormulaResultType());
            case Cell.CELL_TYPE_BLANK:
                return StringUtils.EMPTY;
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                final BigDecimal v = new BigDecimal(cell.getNumericCellValue());
                return v.round(new MathContext(4, RoundingMode.HALF_UP));
            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue().getString();
//            case HSSFCell.CELL_TYPE_ERROR: ("#Error");
//            default: ("Invalid or unknown cell value type!");
        }
        return null;
    }

    public static String getTypeName(int valueType) {
        switch (valueType) {
            case Cell.CELL_TYPE_BLANK:
                return "Empty";
            case Cell.CELL_TYPE_BOOLEAN:
                return "Boolean";
            case Cell.CELL_TYPE_NUMERIC:
                return "Numeric";
            case Cell.CELL_TYPE_STRING:
                return "String";
            case Cell.CELL_TYPE_ERROR:
                return "Error";
            default:
                return "Unknown";
        }
    }

    public static final Logger log = LoggerFactory.getLogger(HSSF2PDFUtils.class);
}
