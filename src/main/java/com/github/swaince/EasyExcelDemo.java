package com.github.swaince;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.extern.slf4j.Slf4j;

import javax.xml.ws.soap.Addressing;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Objects;

/**
 * @author zhangth
 * @date 2019/12/20 22:15
 * @email zhangth@cnegroup.com
 * @description
 */
@Slf4j
public class EasyExcelDemo {

    static String source = "source.xlsx";

    public static void main(String[] args) {

        // 默认使用template模式
        if(read2()) {
            // 若没有提供模板，则需要使用COPY AND WRITE 模式
            read();
        };

    }

    public static void read() {

        ExcelReadAndWriteEventListener listener = new ExcelReadAndWriteEventListener(source);

        ExcelReader excelReader = EasyExcelFactory.read(source, listener)
                .headRowNumber(0)
                .doReadAll();
        listener.closeAll();
        excelReader.finish();

    }

    public static boolean read2() {

        WriteWithTemplateEventListener listener = new WriteWithTemplateEventListener(source);

        ExcelReader excelReader = EasyExcelFactory.read(source, listener)
                .doReadAll();
        listener.closeAll();
        excelReader.finish();
        return listener.isNotTempate();
    }

}
