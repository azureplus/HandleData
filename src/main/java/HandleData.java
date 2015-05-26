import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;


public class HandleData
{
    static Logger logger = Logger.getLogger(HandleData.class);
    static Path path;
    static String[] title;
    static String sortColumn;
    static String sortColumn1;
    static List<Map<String, String>> sortedData;
    static List<Map<String, String>> sortedData1;
    static Map<String, Map<String, String>> resultData;
    static String resultPathStr;

    public static void main(String[] args) throws Exception
    {
        path = Paths.get(args[0]);
        resultPathStr = path.toRealPath().toString().replace(".xls", "_result.xls");
        sortColumn = args[1];
        sortColumn1 = args[2];
        logger.info(path + "," + path.toFile().exists() + "," + sortColumn + "," + sortColumn1);
        InputStream inputStream = new FileInputStream(path.toRealPath().toString());
        title = readExcelTitle(inputStream);
        readExcelContent();
        sortedData = new ArrayList<>();
        sortData(sortedData, sortColumn);
        filterData(sortedData, sortColumn1);
        //  printData(sortedData, sortColumn);
        sortedData1 = new ArrayList<>();
        sortData(sortedData1, sortColumn1);
        filterData(sortedData1, sortColumn);
        // printData(sortedData1, sortColumn1);
        resultData = new HashMap<>();
        logger.info("开始mergeData...");
        mergeData(resultData, sortedData);
        mergeData(resultData, sortedData1);
        logger.info("完成mergeData...共计：" + resultData.size() + "行");
        writeResultData();
    }

    private static POIFSFileSystem fs;
    private static HSSFWorkbook wb;
    private static HSSFSheet sheet;
    private static HSSFRow row;

    public static String[] readExcelTitle(InputStream is) throws IOException
    {
        try
        {
            fs = new POIFSFileSystem(is);
            wb = new HSSFWorkbook(fs);
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
        finally
        {
            wb.close();
        }
        sheet = wb.getSheetAt(0);
        row = sheet.getRow(0);
        int colNum = row.getPhysicalNumberOfCells();
        System.out.println("colNum:" + colNum);
        String[] title = new String[colNum];
        for (int i = 0; i < colNum; i++)
        {
            title[i] = getStringCellValue(row.getCell(i));
        }
        logger.info(Arrays.toString(title));
        return title;
    }

    private static List<Map<String, String>> data;
    static int rowSize;
    static int columnSize;

    private static void readExcelContent()
    {
        rowSize = sheet.getPhysicalNumberOfRows();
        columnSize = title.length;
        logger.info("共计行数：" + rowSize + " 列数:" + columnSize);
        data = new ArrayList<>(rowSize);
        int i = 1;
        int j = 0;
        long t = System.currentTimeMillis();
        try
        {
            for (i = 1; i < rowSize; i++)
            {
                row = sheet.getRow(i);
                Map<String, String> rowContent = new HashMap<>(columnSize);
                for (j = 0; j < columnSize; j++)
                {
                    rowContent.put(title[j], getStringCellValue(row.getCell(j)));
                }
                rowContent.put("uid", String.valueOf(i));//添加行ID
                data.add(rowContent);
            }
        }
        catch (Exception e)
        {
            logger.info(i + "," + j);
            e.printStackTrace();
        }
        logger.info("耗时：" + (System.currentTimeMillis() - t) / 1000.0 + "秒");
        //logger.debug(Arrays.toString(data.toArray()));
    }

    static void printData(List<Map<String, String>> d, String cn)
    {
        for (int i = 0; i < d.size(); i++)
        {
            logger.info(d.get(i).get(cn));
        }
    }

    /**
     * 按照列来排序
     *
     * @param columnName
     */
    static void sortData(List<Map<String, String>> sdata, String columnName)
    {
        logger.info("开始排序");
        long t = System.currentTimeMillis();
        int j = 0;
        for (int i = 0; i < rowSize - 1; i++)
        {
            if (i % 1000 == 0 || i == rowSize - 2)
            {
                logger.info("sort:" + (i + 1) + ",total:" + (rowSize - 1) + ",per: " + (int) ((i + 1) / (1.0 * (rowSize - 1)) * 10000) / 100.0 + "%,耗时：" + (System.currentTimeMillis() - t) / 1000.0 + "秒");
            }
            Map<String, String> v = data.get(i);
            Double compValue = parseDouble(v.get(columnName));
            if (i > 0)
                j = findIndex(sdata, columnName, compValue, 0, i);
            sdata.add(j, v);
            j++;
        }
        logger.info("排序完成：" + columnName + " ,耗时：" + (System.currentTimeMillis() - t) / 1000.0 + "秒");
    }

    static void filterData(List<Map<String, String>> sdata, String columnName)
    {
        long t = System.currentTimeMillis();
        logger.info("开始过滤排序");
        double d = parseDouble(sdata.get(0).get(columnName));
        double c;
        for (int i = 1; i < sdata.size(); i++)
        {
            c = parseDouble(sdata.get(i).get(columnName));
            if (c > d)
            {
                d = c;
            }
            else
            {
                sdata.remove(i);
                i--;
            }
        }
        logger.info("过滤排序完成：" + columnName + " ,耗时：" + (System.currentTimeMillis() - t) / 1000.0 + "秒，" + sdata.size() + "行");
    }

    static void mergeData(Map<String, Map<String, String>> resultData, List<Map<String, String>> mergeList)
    {
        for (int i = 0; i < mergeList.size(); i++)
        {
            Map<String, String> item = mergeList.get(i);
            resultData.put(item.get("uid"), item);
        }
    }

    static int findIndex(List<Map<String, String>> sdata, String columnName, Double insertValue, int start, int end)
    {
        int index = 0;
        int baseIndex = (start + end) / 2;
        if (baseIndex >= end)
            return baseIndex;
        Map<String, String> v = sdata.get(baseIndex);
        Double vc = parseDouble(v.get(columnName));
        if (insertValue < vc)
        {
            index = findIndex(sdata, columnName, insertValue, start, baseIndex);
        }
        else
        {
            index = findIndex(sdata, columnName, insertValue, baseIndex + 1, end);
        }
        return index;
    }

    static double parseDouble(String value)
    {
        if (value == null || value == "")
            return 0.0;
        return Double.parseDouble(value);
    }

    static DecimalFormat df = new DecimalFormat("0");

    /**
     * 获取单元格内容
     *
     * @param cell
     * @return
     */
    private static String getStringCellValue(HSSFCell cell)
    {
        String strCell = "";
        if (cell == null)
        {
            return "";
        }
        switch (cell.getCellType())
        {
            case HSSFCell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                strCell = String.valueOf(df.format(cell.getNumericCellValue()));
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell.equals("") || strCell == null)
        {
            return "";
        }
        return strCell;
    }


    private String getCellFormatValue(HSSFCell cell)
    {
        String cellvalue = "";
        if (cell != null)
        {
            switch (cell.getCellType())
            {
                // 如果当前Cell的Type为NUMERIC
                case HSSFCell.CELL_TYPE_NUMERIC:
                case HSSFCell.CELL_TYPE_FORMULA:
                {
                    // 判断当前的cell是否为Date
                    if (HSSFDateUtil.isCellDateFormatted(cell))
                    {
                        // 如果是Date类型则，转化为Data格式

                        //方法1：这样子的data格式是带时分秒的：2011-10-12 0:00:00
                        //cellvalue = cell.getDateCellValue().toLocaleString();

                        //方法2：这样子的data格式是不带带时分秒的：2011-10-12
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);

                    }
                    // 如果是纯数字
                    else
                    {
                        // 取得当前Cell的数值
                        cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                // 如果当前Cell的Type为STRIN
                case HSSFCell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                // 默认的Cell值
                default:
                    cellvalue = " ";
            }
        }
        else
        {
            cellvalue = "";
        }
        return cellvalue;

    }

    static void writeResultData() throws Exception
    {
        try
        {
            HSSFWorkbook workbook = new HSSFWorkbook();                        // 创建工作簿对象
            FileOutputStream fos = new FileOutputStream(resultPathStr);        // 创建.xls文件
            HSSFSheet sheet = workbook.createSheet();                        // 创建工作表

//            sheet.setDefaultColumnWidth(30);                        // 设置工作表列宽
//            sheet.setDefaultRowHeight((short) 10);                            // 设置工作表行高

            HSSFCellStyle columnTopStyle = getColumnTopStyle(workbook);//获取列头样式对象
            HSSFCellStyle style = getStyle(workbook);                    //单元格样式对象
            //设置列头
            HSSFRow row1 = sheet.createRow((short) 0);                // 在索引0的位置创建行(最顶端的行)
            HSSFCell cell1 = null;                                    // 在索引0的位置创建单元格(左上端)
            // 将列头设置到sheet的单元格中
            for (int i = 0; i < columnSize; i++)
            {
                cell1 = row1.createCell(i);                //创建列头对应个数的单元格
                cell1.setCellType(HSSFCell.CELL_TYPE_STRING);        //设置列头单元格的数据类型
                cell1.setCellValue(title[i]);                        //设置列头单元格的值
               // cell1.setCellStyle(columnTopStyle);                    //设置列头单元格样式
            }
            int i = 1;
            for (Map.Entry<String, Map<String, String>> entry : resultData.entrySet())
            {
                HSSFRow row = sheet.createRow(i); 
                i++;
                Map<String, String> item = entry.getValue();
                for (int j = 0; j < columnSize; j++)
                {
                    HSSFCell cell = row.createCell(j, HSSFCell.CELL_TYPE_STRING);//设置单元格的数据类型
                    cell.setCellValue(item.get(title[j]));                                    //设置单元格的值
                    //cell.setCellStyle(style);                                    //设置单元格样式
                }
            }
            workbook.write(fos);// 将workbook对象输出到文件test.xls
            fos.flush();        // 缓冲
            fos.close();        // 关闭流
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        logger.info("保存完成，保存位置："+resultPathStr);
    }

    /* 
     * 列头单元格样式
     */
    static HSSFCellStyle getColumnTopStyle(HSSFWorkbook workbook)
    {

        // 设置字体
        HSSFFont font = workbook.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 11);
        //字体加粗
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //设置字体名字 
        font.setFontName("Courier New");
        //设置样式; 
        HSSFCellStyle style = workbook.createCellStyle();
        //设置底边框; 
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        //设置底边框颜色;  
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        //设置左边框;   
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        //设置左边框颜色; 
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        //设置右边框; 
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        //设置右边框颜色; 
        style.setRightBorderColor(HSSFColor.BLACK.index);
        //设置顶边框; 
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        //设置顶边框颜色;  
        style.setTopBorderColor(HSSFColor.BLACK.index);
        //在样式用应用设置的字体;  
        style.setFont(font);
        //设置自动换行; 
        style.setWrapText(false);
        //设置水平对齐的样式为居中对齐;  
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置垂直对齐的样式为居中对齐; 
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        return style;

    }

    /*  
   * 列数据信息单元格样式
   */
    static HSSFCellStyle getStyle(HSSFWorkbook workbook)
    {
        // 设置字体
        HSSFFont font = workbook.createFont();
        //设置字体大小
        //font.setFontHeightInPoints((short)10);
        //字体加粗
        //font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //设置字体名字 
        font.setFontName("Courier New");
        //设置样式; 
        HSSFCellStyle style = workbook.createCellStyle();
        //设置底边框; 
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        //设置底边框颜色;  
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        //设置左边框;   
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        //设置左边框颜色; 
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        //设置右边框; 
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        //设置右边框颜色; 
        style.setRightBorderColor(HSSFColor.BLACK.index);
        //设置顶边框; 
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        //设置顶边框颜色;  
        style.setTopBorderColor(HSSFColor.BLACK.index);
        //在样式用应用设置的字体;  
        style.setFont(font);
        //设置自动换行; 
        style.setWrapText(false);
        //设置水平对齐的样式为居中对齐;  
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置垂直对齐的样式为居中对齐; 
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        return style;

    }
}
