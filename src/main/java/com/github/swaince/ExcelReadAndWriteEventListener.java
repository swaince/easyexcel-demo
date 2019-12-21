package com.github.swaince;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.extern.slf4j.Slf4j;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author zhangth
 * @date 2019/12/21 10:39
 * @email zhangth@cnegroup.com
 * @description
 */
@Slf4j
public class ExcelReadAndWriteEventListener<T> extends AnalysisEventListener<T> {

    public ExcelReadAndWriteEventListener(String source) {
        this.source = source;
        excelWriter = EasyExcel.write(System.currentTimeMillis() + source)
                .useDefaultStyle(Boolean.FALSE)
                .build();
    }

    private String source;
    private String sheetName;
    private Integer sheetNo;
    private boolean isInitSheet;
    private LinkedHashMap<Integer, Object> head;
    private List<List<String>> headers;
    private List<List<Object>> dataList = new ArrayList<>();
    private ExcelWriter excelWriter;
    private List<String> handleList = Arrays.asList("遥测", "遥信", "电度");
    private static final Integer DEVICE_CODE_INDEX = 4;
    private static final Integer DEVICE_NAME_INDEX = 3;
    private AtomicInteger count;
    private static final int ATOMIC_INIT_SIZE = 1;

    public void closeAll() {
        excelWriter.finish();
        log.info("文件流已经正常关闭");
    }

    @Override
    public void invoke(T data, AnalysisContext context) {

        Integer rowIndex = context.readRowHolder().getRowIndex();

        if (!isInitSheet) {
            init(context);
            head = (LinkedHashMap<Integer, Object>) data;
            headers = new ArrayList<>();
            for (Object value : head.values()) {
                List<String> header = new ArrayList<>();
                if (Objects.nonNull(value)) {
                    header.add(value.toString());
                }
                headers.add(header);
            }
            log.info("SheetNo[ {} ], SheetName[ {} ]信息已经初始化, Headers {} ",
                    sheetNo, sheetName, headers);
            // 处理完header后就不在继续
            return;
        }
        if (Objects.nonNull(data)) {
            LinkedHashMap<Integer, Object> dataMap = (LinkedHashMap<Integer, Object>) data;
            // 一行数据
            List<Object> rowData = new ArrayList<>();
            if (handleList.contains(sheetName)) {
                dataMap.put(DEVICE_CODE_INDEX, getDeviceCode(dataMap.get(DEVICE_NAME_INDEX)));
            }
            for (Object value : dataMap.values()) {
                rowData.add(value);
            }

            dataList.add(rowData);
        }
        log.info("SheetNo[ {} ], SheetName[ {} ], RowIndex[ {} ], data[ {} ]",
                sheetNo, sheetName, rowIndex, data);


    }

    /**
     * 模拟查询设备编码
     *
     * @param deviceName
     * @return
     */
    private Object getDeviceCode(Object deviceName) {

        return "" + count.getAndIncrement();
    }

    private void init(AnalysisContext context) {
        ReadSheet readSheet = context.readSheetHolder().getReadSheet();
        sheetName = readSheet.getSheetName();
        sheetNo = readSheet.getSheetNo();
        count = new AtomicInteger(ATOMIC_INIT_SIZE);
        isInitSheet = true;

    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

        WriteSheet writeSheet = EasyExcel.writerSheet(sheetNo, sheetName).build();
        writeSheet.setHead(headers);
        excelWriter.write(dataList, writeSheet);
        clear();
    }

    private void clear() {
        log.info("SheetNo[ {} ], SheetName[ {} ]已经处理完毕，初始化信息开始重置", sheetNo, sheetName);
        sheetName = null;
        sheetNo = null;
        headers = null;
        count = new AtomicInteger(ATOMIC_INIT_SIZE);
        dataList.clear();
        isInitSheet = false;
        log.info("信息重置完成");
    }
}
