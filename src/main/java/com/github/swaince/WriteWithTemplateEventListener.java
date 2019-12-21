package com.github.swaince;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author zhangth
 * @date 2019/12/21 11:35
 * @email zhangth@cnegroup.com
 * @description
 */
@Slf4j
public class WriteWithTemplateEventListener extends AnalysisEventListener {

    public WriteWithTemplateEventListener(String source) {
        this.target = System.currentTimeMillis() + source;
        excelWriter = EasyExcel.write(target)
                .withTemplate(source)
                .build();
    }

    private String target;
    private Boolean isNotTempate;
    private String sheetName;
    private Integer sheetNo;
    private boolean isInitSheet;
    private List<DeviceInfo> dataList = new ArrayList<>();
    private ExcelWriter excelWriter;
    private List<String> handleList = Arrays.asList("遥测", "遥信", "电度");
    private static final Integer DEVICE_CODE_INDEX = 4;
    private static final Integer DEVICE_NAME_INDEX = 3;
    private AtomicInteger count;
    private static final int ATOMIC_INIT_SIZE = 1;

    public void closeAll() {
        excelWriter.finish();
        if (isNotTempate) {
            log.info("临时文件[ {} ]已经删除", target);
            File file = new File(target);
            file.delete();
        }
        log.info("文件流已经正常关闭");
    }

    public boolean isNotTempate() {
        return isNotTempate;
    }

    @Override
    public void invoke(Object data, AnalysisContext context) {

        if (Objects.nonNull(isNotTempate) && isNotTempate) {
            return;
        }
        Integer rowIndex = context.readRowHolder().getRowIndex();

        if (!isInitSheet) {
            init(context);
            log.info("SheetNo[ {} ], SheetName[ {} ]信息已经初始化",
                    sheetNo, sheetName);
        }
        if (Objects.nonNull(data)) {
            LinkedHashMap<Integer, Object> dataMap = (LinkedHashMap<Integer, Object>) data;
            if (handleList.contains(sheetName)&& Objects.nonNull(dataMap.get(DEVICE_NAME_INDEX))) {
                Object deviceCodeTemplate = dataMap.get(DEVICE_CODE_INDEX);
                if (Objects.isNull(deviceCodeTemplate) || !deviceCodeTemplate.toString().startsWith("{.")) {
                    if (Objects.isNull(isNotTempate)) {
                        log.info("当前文件不是模板文件");
                        isNotTempate = true;
                        // 不是模板文件就不在处理
                        return;
                    }
                } else {
                    isNotTempate = false;
                }
                dataList.add(getDeviceCode(dataMap.get(DEVICE_NAME_INDEX)));
            }
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
    private DeviceInfo getDeviceCode(Object deviceName) {

        return new DeviceInfo("" + count.getAndIncrement());
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
        if (Objects.nonNull(isNotTempate) && isNotTempate) {
            return;
        }
        WriteSheet writeSheet = EasyExcel.writerSheet(sheetNo, sheetName)
                .includeColumnIndexes(Arrays.asList(DEVICE_CODE_INDEX))
                .build();
        excelWriter.fill(dataList, writeSheet);
        clear();
    }

    private void clear() {
        log.info("SheetNo[ {} ], SheetName[ {} ]已经处理完毕，初始化信息开始重置", sheetNo, sheetName);
        sheetName = null;
        sheetNo = null;
        count = new AtomicInteger(ATOMIC_INIT_SIZE);
        dataList.clear();
        isInitSheet = false;
        log.info("信息重置完成");
    }

    @Data
    private static class DeviceInfo {

        public DeviceInfo(String deviceCode) {
            this.deviceCode = deviceCode;
        }

        private String deviceCode;
    }
}
