package per.dyp.webcollection;

import per.dyp.webcollection.common.DateUtils;

public class Main {
    public static void main(String args[]) {

        System.out.println(DateUtils.getCurrentTime(DateUtils.HH_mm_ss) + "[开始采集]");
        Collector collector = new QiuTanCollector();
        Object object = collector.collect();
        System.out.println(DateUtils.getCurrentTime(DateUtils.HH_mm_ss) + "[采集完成]");
        System.out.println(DateUtils.getCurrentTime(DateUtils.HH_mm_ss) + "[开始导出]");
        Exporter exporter = new QiuTanExporter();
        exporter.export(object, args[0]);
        System.out.println(DateUtils.getCurrentTime(DateUtils.HH_mm_ss) + "[导出完成]");
    }
}
