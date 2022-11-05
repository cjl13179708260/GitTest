package main.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;

import java.io.BufferedOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class TestEasyExcel {

    public static void getExcel(Head head, List<Dto> bodyList) throws IOException {
        ExcelWriter excelWriter = null;


        try {
            excelWriter = EasyExcel.write("D:/sheet", Dto.class).build();
            for (int i = 0; i < 5; i++) {
                WriteSheet writeSheet = EasyExcel.writerSheet(i,"模板"+i).head(Dto.class).build();
                excelWriter.write(bodyList,writeSheet);
            }
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    public static void main(String[] args) throws IOException {
       Head head = new Head();
       head.setD1("d1");
       head.setD2("d2");
       head.setD3("d3");
       head.setD4("d4");
       head.setD5("d5");
       head.setD6("d6");
       head.setD7("d7");
       head.setD8("d8");
       head.setD9("d9");
       head.setD10("d10");
        Dto dto1 = new Dto();
        dto1.setData1("data111");
        dto1.setData2("data222");
        dto1.setData3("data333");
        dto1.setData4("data444");
        dto1.setData5("data555");
        dto1.setData6("data666");
        dto1.setData7("data777");
        dto1.setData8("data888");
        dto1.setData9("data999");
        dto1.setData10("data100");

        List<Dto> bodyList = new ArrayList<>();
        bodyList.add(dto1);
        bodyList.add(dto1);
        bodyList.add(dto1);
        bodyList.add(dto1);
        bodyList.add(dto1);
        bodyList.add(dto1);
        bodyList.add(dto1);
        bodyList.add(dto1);
        bodyList.add(dto1);
        bodyList.add(dto1);



        getExcel(head, bodyList);

    }

}
