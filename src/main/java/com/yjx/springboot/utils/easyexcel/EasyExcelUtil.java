package com.yjx.springboot.utils.easyexcel;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 *
 */
public class EasyExcelUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(EasyExcelUtil.class);

    /**
     * web导出excel
     *
     * @param response
     * @param fileName
     * @param data
     * @param rowModel
     * @param <T>
     * @throws IOException
     */
    public static <T extends BaseRowModel> void excelExport(HttpServletResponse response, String fileName, List<T> data, Class<T> rowModel) throws IOException {
        setResponse(response, fileName);
        EasyExcelUtil.writeOneSheet(response.getOutputStream(), fileName, data, rowModel);
    }

    private static void setResponse(HttpServletResponse response, String fileName) throws UnsupportedEncodingException {
        // 设置response的Header
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8") + ".xlsx");
    }

    /**
     * 文件名是否合法
     *
     * @param fileName
     * @return
     */
    public static boolean checkFileName(String fileName) {
        if (null == fileName) {
            return false;
        }
        if (!fileName.toLowerCase().endsWith(ExcelTypeEnum.XLS.getValue()) && !fileName.toLowerCase().endsWith(ExcelTypeEnum.XLSX.getValue())) {
            return false;
        }
        return true;
    }

    /**
     * 读取 某sheet 从某行开始
     *
     * @param filePath
     * @param sheetIndex
     * @param startRowIndex
     * @param trim          是否对内容做trim()增加容错
     * @return Object => List<String>
     * @throws Exception
     */
    public static List<Object> readRow(String filePath, Integer sheetIndex, Integer startRowIndex, boolean trim) throws Exception {

        if (!checkFileName(filePath)) {
            throw new Exception("文件格式不合法");
        }

        InputStream in = null;
        try {
            in = new FileInputStream(filePath);
            return EasyExcelUtil.readRow(in, sheetIndex, startRowIndex, trim);
        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                }
            }
        }

        return null;
    }

    /**
     * 读取 某sheet 从某行开始
     *
     * @param in
     * @param sheetIndex
     * @param startRowIndex
     * @param trim          是否对内容做trim()增加容错
     * @return Object => List<String>
     * @throws Exception
     */
    public static List<Object> readRow(InputStream in, Integer sheetIndex, Integer startRowIndex, boolean trim) throws Exception {

        try {
            final List<Object> rows = new ArrayList<Object>();
            Sheet sheet = new Sheet(sheetIndex, startRowIndex);  // (某sheet, 某行)
            new ExcelReader(in, null, new AnalysisEventListener() {  //ExcelListener获取解析结果 并操作 
                @Override
                public void invoke(Object object, AnalysisContext context) {
                    if (null != object) {
                        rows.add(object);
                    }
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    // rows.clear(); 
                }
            }, trim).read(sheet);

            return rows;

        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                }
            }
        }

        return null;
    }

    /**
     * 读取 指定sheet 指定列 文本格式
     *
     * @param filePath
     * @param sheetIndex
     * @param columnIndex
     * @param startRowIndex
     * @return
     * @throws Exception
     */
    public static List<String> readOneSheetOneColumn(String filePath, int sheetIndex, int startRowIndex, int columnIndex) throws Exception {

        if (!checkFileName(filePath)) {
            throw new Exception("文件格式不合法");
        }

        InputStream in = null;
        try {
            in = new FileInputStream(filePath);
            return EasyExcelUtil.readOneSheetOneColumn(in, sheetIndex, startRowIndex, columnIndex);

        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                }
            }
        }
        return null;
    }


    /**
     * 读取 指定sheet 指定列 文本格式
     *
     * @param in
     * @param sheetIndex
     * @param columnIndex
     * @param startRowIndex
     * @return
     * @throws Exception
     */
    public static List<String> readOneSheetOneColumn(InputStream in, final int sheetIndex, final int startRowIndex, final int columnIndex) throws Exception {

        final List<String> rows = new ArrayList<String>();
        try {
            Sheet sheet = new Sheet(sheetIndex, startRowIndex);
            new ExcelReader(in, null, new AnalysisEventListener() {
                @Override
                public void invoke(Object object, AnalysisContext context) {
                    if (null != object) {
                        List row = (List<String>) object; // 未指定模型 默认每行为 List<String>
                        if (row.size() > columnIndex) {
                            rows.add((String) row.get(columnIndex));
                        }
                    }
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    // rows.clear(); 
                }
            }, false).read(sheet);

            return rows;

        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                }
            }
        }

        return null;
    }


    /**
     * 读取 某 sheet 从某行开始
     *
     * @param filePath
     * @param sheetIndex
     * @param startRowIndex
     * @param listener      外部监听类
     * @param rowModel
     * @return
     */
    public static <T extends BaseRowModel> List<T> readRow(String filePath, Integer sheetIndex, Integer startRowIndex, AbstractListener<T> listener, Class<T> rowModel) {
        InputStream in = null;
        try {
            in = new FileInputStream(filePath);
            return EasyExcelUtil.readRow(in, sheetIndex, startRowIndex, listener, rowModel);
        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                }
            }
        }

        return null;
    }


    /**
     * 读取 某 sheet 从某行开始
     *
     * @param in
     * @param sheetIndex
     * @param startRowIndex
     * @param listener
     * @param rowModel
     * @return
     */
    public static <T extends BaseRowModel> List<T> readRow(InputStream in, Integer sheetIndex, Integer startRowIndex, AbstractListener<T> listener, Class<T> rowModel) {
        try {
            Sheet sheet = new Sheet(sheetIndex, startRowIndex, rowModel);
            ExcelReader reader = new ExcelReader(in, null, listener, false);
            reader.read(sheet);

            return listener.getDataList();
        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                }
            }
        }

        return null;
    }


    /**
     * 读取 所有 sheet
     *
     * @param in
     * @param listener
     * @param rowModel
     * @param listener
     * @param rowModel
     * @return
     */
    public static <T extends BaseRowModel> List<T> readAllSheetRow(InputStream in, AbstractListener<T> listener, Class<T> rowModel) {
        try {
            ExcelReader reader = new ExcelReader(in, null, listener, false);
            for (Sheet sheet : reader.getSheets()) {
                sheet.setClazz(rowModel);
                reader.read(sheet);
            }
            return listener.getDataList();
        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                }
            }
        }

        return null;
    }

    /**
     * 读取 所有 sheet
     *
     * @param filePath
     * @param listener
     * @param rowModel
     * @return
     */
    public static <T extends BaseRowModel> List<T> readAllSheetRow(String filePath, AbstractListener<T> listener, Class<T> rowModel) {
        InputStream in = null;
        try {
            in = new FileInputStream(filePath);
            return EasyExcelUtil.readAllSheetRow(in, listener, rowModel);
        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                }
            }
        }

        return null;

    }

    /**
     * 写xlsx 单个Sheet
     *
     * @param out
     * @param fileName
     * @param headers  表头
     * @param data
     */
    public static void writeOneSheet(OutputStream out, String fileName, List<String> headers, List<List<Object>> data) {
        Sheet sheet = new Sheet(1, 0);
        sheet.setSheetName(fileName);

        if (null != headers) {
            List<List<String>> list = new ArrayList<List<String>>();
            headers.forEach(e -> list.add(Collections.singletonList(e)));
            sheet.setHead(list);
        }

        ExcelWriter writer = null;
        try {
            writer = EasyExcelFactory.getWriter(out);  // xlsx
            writer.write1(data, sheet);
            writer.finish();
        } catch (Exception e2) {
            LOGGER.error("EasyExcelUtil writer error: {}", e2);
        } finally {
            if (null != out) {
                try {
                    out.close();
                } catch (IOException e) {

                }
            }
        }

    }

    /**
     * 写xlsx 单个Sheet
     *
     * @param filePath
     * @param fileName
     * @param headers
     * @param data
     * @throws Exception
     */
    public static void writeOneSheet(String filePath, String fileName, List<String> headers, List<List<Object>> data) throws Exception {
        if (!checkFileName(filePath)) {
            throw new Exception("文件格式不合法");
        }

        OutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            EasyExcelUtil.writeOneSheet(out, fileName, headers, data);
        } catch (FileNotFoundException e) {
            LOGGER.error("EasyExcelUtil writer error: {}", e);
        } finally {
            if (null != out) {
                try {
                    out.close();
                } catch (IOException e) {
                }
            }
        }
    }

    /**
     * 写xlsx 单个Sheet
     *
     * @param out
     * @param fileName
     * @param data
     * @param rowModel T
     */
    public static <T extends BaseRowModel> void writeOneSheet(OutputStream out, String fileName, List<T> data, Class<T> rowModel) {
        Sheet sheet = new Sheet(1, 0, rowModel);
        sheet.setSheetName(fileName);
        sheet.setAutoWidth(true);

        ExcelWriter writer = null;
        try {
            writer = EasyExcelFactory.getWriter(out);  // xlsx
            writer.write(data, sheet);
            writer.finish();
        } catch (Exception e) {
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != out) {
                try {
                    out.close();
                } catch (IOException e) {
                }
            }
        }
    }


    /**
     * 写xlsx 单个Sheet
     *
     * @param filePath
     * @param fileName
     * @param data
     * @param rowModel
     * @throws Exception
     */
    public static <T extends BaseRowModel> void writeOneSheet(String filePath, String fileName, List<T> data, Class<T> rowModel) throws Exception {
        if (!checkFileName(filePath)) {
            throw new Exception("文件格式不合法");
        }

        OutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            EasyExcelUtil.writeOneSheet(out, fileName, data, rowModel);
        } catch (FileNotFoundException e) {
            LOGGER.error("EasyExcelUtil writer error: {}", e);
        } finally {
            if (null != out) {
                try {
                    out.close();
                } catch (IOException e) {
                }
            }
        }

    }

    /**
     * 写xlsx 多个个Sheet
     *
     * @param out
     * @param sheetSize
     * @param data
     * @param rowModel
     * @throws Exception
     */
    public static <T extends BaseRowModel> void writeSheets(OutputStream out, Integer sheetSize, List<T> data, Class<T> rowModel) throws Exception {
        int totalSize = data.size();
        if (totalSize <= 0) {
            throw new Exception("data无数据");
        }
        ExcelWriter writer = null;
        try {
            writer = EasyExcelFactory.getWriter(out);  // xlsx
            int sheetNum = totalSize % sheetSize == 0 ? totalSize / sheetSize : totalSize / sheetSize + 1;

            Sheet sheet = null;
            int fromIndex = 0;
            int toIndex = 0;
            for (int i = 1; i <= sheetNum; i++) {
                fromIndex = (i - 1) * sheetSize;
                toIndex = i * sheetSize - 1;
                toIndex = Math.min(toIndex, totalSize);
                sheet = new Sheet(i, 0, rowModel);
                sheet.setSheetName("sheet" + i);
                writer.write(data.subList(fromIndex, toIndex), sheet);
            }
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
            LOGGER.error("EasyExcelUtil read error: {}", e);
        } finally {
            if (null != out) {
                try {
                    out.close();
                } catch (IOException e) {
                }
            }
        }


    }

    /**
     * 写xlsx 多个个Sheet
     *
     * @param filePath
     * @param sheetSize
     * @param data
     * @param rowModel
     * @throws Exception
     */
    public static <T extends BaseRowModel> void writeSheets(String filePath, Integer sheetSize, List<T> data, Class<T> rowModel) throws Exception {
        if (!checkFileName(filePath)) {
            throw new Exception("文件格式不合法");
        }

        OutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            EasyExcelUtil.writeSheets(out, sheetSize, data, rowModel);
        } catch (FileNotFoundException e) {
            LOGGER.error("EasyExcelUtil writer error: {}", e);
        } finally {
            if (null != out) {
                try {
                    out.close();
                } catch (IOException e) {
                }
            }
        }

    }

}
