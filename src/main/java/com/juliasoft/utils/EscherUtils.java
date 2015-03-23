package com.juliasoft.utils;

import org.apache.poi.ddf.*;
import org.apache.poi.hssf.record.EscherAggregate;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class EscherUtils {

    public static void processImages(HSSFSheet sheet) {
        final EscherAggregate drawingAggregate = sheet.getDrawingEscherAggregate();
        if (drawingAggregate != null) {
            for (final EscherRecord record : drawingAggregate.getEscherRecords()) {
                iterateRecords(record, 1);
            }
        }
    }

    //
    // http://stackoverflow.com/questions/27011634/how-to-get-pictures-with-names-from-an-xls-file-using-apache-poi
    //
    public static void iterateRecords(EscherRecord escherRecord, int level) {
        for (final EscherRecord rec : escherRecord.getChildRecords()) {
            log.info("== Escher type: {}", rec.getClass().getName());
            log.info(" {}", rec);
//            if (childRecord instanceof EscherClientAnchorRecord) {
//                final EscherClientAnchorRecord ar = (EscherClientAnchorRecord) childRecord;
//                log.info("TopLeft: C:{} R:{}, X:{}, Y:{}", ar.getCol1(), ar.getRow1(), ar.getDx1(), ar.getDy1());
//                log.info("BottomRight: C:{} R:{}, X:{}, Y:{}", ar.getCol2(), ar.getRow2(), ar.getDx2(), ar.getDy2());
//            }
//            if (childRecord instanceof EscherClientDataRecord) {
//                final EscherClientDataRecord dr = (EscherClientDataRecord) childRecord;
//                log.info("EscherClientDataRecord Size: {}", dr.getRecordSize());
//            }
//
            if (rec instanceof EscherOptRecord) {
                final EscherOptRecord or = (EscherOptRecord) rec;
                log.info("OptRecord: {}", or);
//                String key = String.valueOf(level);
//                if (shapes.containsKey(key)) {
//                    shapes.get(key).setOptRecord(or);
//                } else {
//                    final ShapeInfo shape = new ShapeInfo(level);
//                    shape.setOptRecord(or);
//                    shapes.put(key, shape);
//                }
            }
            else if (rec instanceof EscherClientAnchorRecord) {
                final EscherClientAnchorRecord ar = (EscherClientAnchorRecord) rec;
                log.info("AnchorRecord: {}", ar);
//                log.info("TopLeft: C:{} R:{}, X:{}, Y:{}", ar.getCol1(), ar.getRow1(), ar.getDx1(), ar.getDy1());
//                log.info("BottomRight: C:{} R:{}, X:{}, Y:{}", ar.getCol2(), ar.getRow2(), ar.getDx2(), ar.getDy2());
//                String key = String.valueOf(level);
//                if (shapes.containsKey(key)) {
//                    shapes.get(key).setAnchorRecord(ar);
//                } else {
//                    final ShapeInfo shape = new ShapeInfo(level);
//                    shape.setAnchorRecord(ar);
//                    shapes.put(key, shape);
//                }
            }

            if (rec.getChildRecords().size() > 0) {
                iterateRecords(rec, ++level);
            }
        }
    }

    private static final Logger log = LoggerFactory.getLogger(EscherUtils.class);
}