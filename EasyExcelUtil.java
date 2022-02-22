package com.example.demoaac.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.alibaba.excel.write.style.row.SimpleRowHeightStyleStrategy;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author aac
 * @version jdk1.8
 * @since 2022/2/16
 */
public class EasyExcelUtil<T> {

    /**
     * 导出excel -对象
     * @param url      导出路径
     * @param fileName 文件名
     * @param clazz    类
     * @param list     list数据
     * @param <T>      泛型
     */
    public static <T> void exportExcel(String url, String fileName, Class<T> clazz, List<T> list, Map<String, Object> map) {
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = setCellStyle(map);
        //.class就是查询出来的list里的泛型
        EasyExcel.write(url + fileName + ExcelTypeEnum.XLSX.getValue(), clazz)
                //.registerWriteHandler(new CustomSheetWriteHandler()) // 自定义拦截器.对第一列第一行和第二行的数据新增下拉框，显示 测试1 测试2
                //.registerWriteHandler(new CustomCellWriteHandler()) //自定义拦截器。对第一行第一列的头超链接到:https://github.com/alibaba/easyexcel
                .registerWriteHandler(new SimpleColumnWidthStyleStrategy(25)) //简单的列宽策略
                .registerWriteHandler(new SimpleRowHeightStyleStrategy((short) 25, (short) 25)) //简单的行高策略 头行高  列行高
                //.registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()) //自动列宽 不够精确
                .registerWriteHandler(horizontalCellStyleStrategy)
                .autoCloseStream(Boolean.FALSE).sheet("sheet1")
                //list就是查询出来的list泛型加了注解的domain
                .doWrite(list);
    }

    /**
     * 导出excel -流
     * @param response 请求
     * @param fileName 文件名
     * @param clazz    类
     * @param list     list数据
     * @param map      样式设置
     * @param <T>      泛型
     */
    public static <T> void exportExcel(HttpServletResponse response, String fileName, Class<T> clazz, List<T> list, Map<String, Object> map) {
        setResponseParam(response, fileName);
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = setCellStyle(map);
        //.class就是查询出来的list里的泛型
        try {
            EasyExcel.write(response.getOutputStream(), clazz)
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .registerWriteHandler(new SimpleColumnWidthStyleStrategy(25)) //简单的列宽策略
                    .registerWriteHandler(new SimpleRowHeightStyleStrategy((short) 25, (short) 25)) //简单的行高策略 头行高  列行高
                    .autoCloseStream(Boolean.FALSE).sheet("sheet1")
                    //list就是查询出来的list泛型加了注解的domain
                    .doWrite(list);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 导出excel 自定义表头
     * @param url      导出路径
     * @param fileName 文件名
     * @param heads    表头
     * @param list     数据
     * @param map      样式设置
     */
    public static void exportExcel(String url, String fileName, List<List<String>> heads, List<List<Object>> list, Map<String, Object> map) {
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = setCellStyle(map);
        //.class就是查询出来的list里的泛型
        EasyExcel.write(url + fileName + ExcelTypeEnum.XLSX.getValue())
                .head(heads)
                .registerWriteHandler(horizontalCellStyleStrategy)
                .autoCloseStream(Boolean.FALSE).sheet("sheet1")
                //list就是查询出来的list泛型加了注解的domain
                .doWrite(list);
    }

    /**
     * 导出excel 自定义表头 流
     * @param response 请求
     * @param fileName 文件名
     * @param heads    表头
     * @param list     数据
     * @param map      样式设置
     */
    public static void exportExcel(HttpServletResponse response, String fileName, List<List<String>> heads, List<List<Object>> list, Map<String, Object> map) {
        setResponseParam(response, fileName);
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = setCellStyle(map);
        //.class就是查询出来的list里的泛型
        try {
            EasyExcel.write(response.getOutputStream())
                    .head(heads)
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .autoCloseStream(Boolean.FALSE).sheet("sheet1")
                    //list就是查询出来的list泛型加了注解的domain
                    .doWrite(list);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 简单的模板填充
     * @param templateUrl 模板路径
     * @param fileUrl     导出路径
     * @param fileName    导出文件名
     * @param t           对象
     * @param <T>         泛型
     */
    public static <T> void simpleFillByObject(String templateUrl, String fileUrl, String fileName, T t) {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // 方案1 根据对象填充
        String url = fileUrl + fileName + ExcelTypeEnum.XLSX.getValue();
        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        EasyExcel.write(url).withTemplate(templateUrl).sheet().doFill(t);
    }

    /**
     * 简单的模板填充 -流
     * @param response 请求
     * @param fileName 文件名
     * @param file     模板文件
     * @param t        对象
     * @param <T>      泛型
     */
    public static <T> void simpleFillByObject(HttpServletResponse response, String fileName, MultipartFile file, T t) {
        setResponseParam(response, fileName);
        try {
            EasyExcel.write(response.getOutputStream()).withTemplate(file.getInputStream()).sheet().doFill(t);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 填充列表
     * @param templateUrl 模板路径
     * @param fileUrl     导出路径
     * @param fileName    导出文件名
     * @param list        数据列表
     * @param <T>         泛型
     */
    public static <T> void listFill(String templateUrl, String fileUrl, String fileName, List<T> list) {
        String url = fileUrl + fileName + ExcelTypeEnum.XLSX.getValue();
        EasyExcel.write(url).withTemplate(templateUrl).sheet().doFill(list);
    }

    /**
     * 填充列表 流
     * @param response 请求
     * @param fileName 文件名
     * @param file     模板文件
     * @param list     数据列表
     * @param <T>      泛型
     */
    public static <T> void listFill(HttpServletResponse response, String fileName, MultipartFile file, List<T> list) {
        setResponseParam(response, fileName);
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // 填充list 的时候还要注意 模板中{.} 多了个点 表示list

        // 方案1 一下子全部放到内存里面 并填充
        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        try {
            EasyExcel.write(response.getOutputStream()).withTemplate(file.getInputStream()).sheet().doFill(list);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 数据量小的复杂填充  数据量大根据情况手动封装
     * @param templateUrl 模板文件路径
     * @param fileUrl     导出路径
     * @param fileName    文件名
     * @param list        列表数据
     * @param map         统计、标题等数据
     * @param <T>         泛型
     */
    public static <T> void complexFill(String templateUrl, String fileUrl, String fileName, List<T> list, Map<String, Object> map) {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // {} 代表普通变量 {.} 代表是list的变量
        String url = fileUrl + fileName + ExcelTypeEnum.XLSX.getValue();
        ExcelWriter excelWriter = EasyExcel.write(url).withTemplate(templateUrl).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // 这里注意 入参用了forceNewRow 代表在写入list的时候不管list下面有没有空行 都会创建一行，然后下面的数据往后移动。默认 是false，会直接使用下一行，如果没有则创建。
        // forceNewRow 如果设置了true,有个缺点 就是他会把所有的数据都放到内存了，所以慎用
        // 简单的说 如果你的模板有list,且list不是最后一行，下面还有数据需要填充 就必须设置 forceNewRow=true 但是这个就会把所有数据放到内存 会很耗内存
        // 如果数据量大 list不是最后一行 参照下一个
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        excelWriter.fill(list, fillConfig, writeSheet);
        excelWriter.fill(map, writeSheet);
        excelWriter.finish();
    }

    /**
     * 数据量小的复杂填充 流
     * @param response 请求
     * @param fileName 文件名
     * @param file     模板文件
     * @param list     列表数据
     * @param map      统计、标题等数据
     * @param <T>      泛型
     */
    public static <T> void complexFill(HttpServletResponse response, String fileName, MultipartFile file, List<T> list, Map<String, Object> map) {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // {} 代表普通变量 {.} 代表是list的变量
        setResponseParam(response, fileName);
        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(file.getInputStream()).build();
        } catch (IOException e) {
            e.printStackTrace();
        }
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // 这里注意 入参用了forceNewRow 代表在写入list的时候不管list下面有没有空行 都会创建一行，然后下面的数据往后移动。默认 是false，会直接使用下一行，如果没有则创建。
        // forceNewRow 如果设置了true,有个缺点 就是他会把所有的数据都放到内存了，所以慎用
        // 简单的说 如果你的模板有list,且list不是最后一行，下面还有数据需要填充 就必须设置 forceNewRow=true 但是这个就会把所有数据放到内存 会很耗内存
        // 如果数据量大 list不是最后一行 参照下一个
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        excelWriter.fill(list, fillConfig, writeSheet);
        excelWriter.fill(map, writeSheet);
        excelWriter.finish();
    }

    /**
     * 横向的填充
     * @param templateUrl 模板文件路径
     * @param fileUrl     导出路径
     * @param fileName    文件名
     * @param list        列表数据
     * @param map         统计、标题等数据
     * @param <T>         泛型
     */
    public static <T> void horizontalFill(String templateUrl, String fileUrl, String fileName, List<T> list, Map<String, Object> map) {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // {} 代表普通变量 {.} 代表是list的变量
        String url = fileUrl + fileName + ExcelTypeEnum.XLSX.getValue();
        ExcelWriter excelWriter = EasyExcel.write(url).withTemplate(templateUrl).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        excelWriter.fill(list, fillConfig, writeSheet);
        //map.put("date", "2019年10月9日13:28:28");
        excelWriter.fill(map, writeSheet);
        // 别忘记关闭流
        excelWriter.finish();
    }

    /**
     * 横向的填充 流
     * @param response 请求
     * @param fileName 文件名
     * @param file     模板文件
     * @param list     列表数据
     * @param map      统计、标题等数据
     * @param <T>      泛型
     */
    public static <T> void horizontalFill(HttpServletResponse response, String fileName, MultipartFile file, List<T> list, Map<String, Object> map) {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // {} 代表普通变量 {.} 代表是list的变量
        setResponseParam(response, fileName);
        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(file.getInputStream()).build();
        } catch (IOException e) {
            e.printStackTrace();
        }
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        excelWriter.fill(list, fillConfig, writeSheet);
        //map.put("date", "2019年10月9日13:28:28");
        excelWriter.fill(map, writeSheet);
        // 别忘记关闭流
        excelWriter.finish();
    }

    /**
     * 读数据示例
     * @param file 文件
     * @param fileUrl 文件路径
     * @param clazz 类
     * @param <T> 泛型
     */
    public static <T> void readExcel(MultipartFile file, String fileUrl, Class<T> clazz) {
        List<T> list = new ArrayList<T>();
        /*
         * EasyExcel 读取 是基于SAX方式
         * 因此在解析时需要传入监听器
         */
        // 写法1：
        // 第一个参数 为 excel文件路径
        // 读取时的数据类型
        // 监听器
        try {
            EasyExcel.read(file.getInputStream(), clazz, new AnalysisEventListener<T>() {

                // 每读取一行就调用该方法
                @Override
                public void invoke(T data, AnalysisContext context) {
                    list.add(data);
                }

                // 全部读取完成就调用该方法
                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    //System.out.println("读取完成");
                }
            })
                    // 这里注意 我们也可以registerConverter来指定自定义转换器， 但是这个转换变成全局了， 所有java为string,excel为string的都会用这个转换器。
                    // 如果就想单个字段使用请使用@ExcelProperty 指定converter
                    // .registerConverter(new CustomStringStringConverter())
                    // 这里可以设置1，因为头就是一行。如果多行头，可以设置其他值。不传入也可以，因为默认会根据DemoData 来解析，他没有指定头，也就是默认1行
                    //.headRowNumber(1)
                    .sheet().doRead();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 写法2：
        ExcelReader excelReader = EasyExcel.read(fileUrl, clazz, new AnalysisEventListener<T>() {

            // 每读取一行就调用该方法
            @Override
            public void invoke(T data, AnalysisContext context) {
                list.add(data);
            }

            // 全部读取完成就调用该方法
            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
                //System.out.println("读取完成");
            }
        }).build();
        ReadSheet readSheet = EasyExcel.readSheet(0).build();
        excelReader.read(readSheet);
        // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
        excelReader.finish();

        //写法3：不创建对象读
        List<Map<Integer, String>> mapList = new ArrayList<Map<Integer, String>>();
        try {
            EasyExcel.read(file.getInputStream(), new AnalysisEventListener<Map<Integer, String>>() {

                // 每读取一行就调用该方法
                @Override
                public void invoke(Map<Integer, String> data, AnalysisContext context) {
                    mapList.add(data);
                }

                // 全部读取完成就调用该方法
                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    //System.out.println("读取完成");
                }
            }).sheet().doRead();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void setResponseParam(HttpServletResponse response, String fileName) {
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ExcelTypeEnum.XLSX.getValue());
        // 响应类型,编码
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }

    private static HorizontalCellStyleStrategy setCellStyle(Map<String, Object> map) {
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景设置为红色
        if (StringUtils.isEmpty(map.get("color"))) {
            headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        } else {
            headWriteCellStyle.setFillForegroundColor((Short) map.get("color"));
        }
        //字体样式设置
        WriteFont headWriteFont = new WriteFont();
        //设置字体大小
        if (StringUtils.isEmpty(map.get("headSize"))) {
            headWriteFont.setFontHeightInPoints((short) 20);
        } else {
            headWriteFont.setFontHeightInPoints((Short) map.get("headSize"));
        }
        //设置字体是否有边框
        if (StringUtils.isEmpty(map.get("headBold"))) {
            headWriteFont.setBold(true);
        } else {
            headWriteFont.setBold((Boolean) map.get("headBold"));
        }
        headWriteCellStyle.setWriteFont(headWriteFont);

        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        WriteFont contentWriteFont = new WriteFont();
        //设置字体大小
        if (StringUtils.isEmpty(map.get("contentSize"))) {
            headWriteFont.setFontHeightInPoints((short) 13);
        } else {
            headWriteFont.setFontHeightInPoints((short) map.get("contentSize"));
        }
        if (StringUtils.isEmpty(map.get("contentBold"))) {
            headWriteFont.setBold(true);
        } else {
            headWriteFont.setBold((Boolean) map.get("contentBold"));
        }
        //设置内容字体
        contentWriteCellStyle.setWriteFont(contentWriteFont);
        //设置 水平居中
        if (StringUtils.isEmpty(map.get("contentCenter"))) {
            contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        } else {
            contentWriteCellStyle.setHorizontalAlignment((HorizontalAlignment) map.get("contentCenter"));
        }
        return new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
    }
}
