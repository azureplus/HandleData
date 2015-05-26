import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
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
    static LinkedList<Map<String, String>> sortedData;
    static LinkedList<Map<String, String>> sortedData1;

    public static void main(String[] args) throws Exception
    {
        path = Paths.get(args[0]);
        sortColumn = args[1];
        sortColumn1 = args[2];
        logger.info(path + "," + path.toFile().exists() + "," + sortColumn + "," + sortColumn1);
        InputStream inputStream = new FileInputStream(path.toRealPath().toString());
        title = readExcelTitle(inputStream);
        readExcelContent();
        sortedData = new LinkedList<>();
        sortData(sortedData, sortColumn);
        printData(sortedData, sortColumn);
        sortedData1 = new LinkedList<>();
        sortData(sortedData1, sortColumn1);
    }

    private static POIFSFileSystem fs;
    private static HSSFWorkbook wb;
    private static HSSFSheet sheet;
    private static HSSFRow row;

    public static String[] readExcelTitle(InputStream is)
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

    private static void printData(List<Map<String, String>> d, String cn)
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
    private static void sortData(LinkedList<Map<String, String>> sdata, String columnName)
    {
        logger.info("开始排序");
        long t = System.currentTimeMillis();
        double baseValue = Double.parseDouble(data.get(0).get(columnName));
        int baseIndex = 0;
        sdata.addFirst(data.get(0));
        Double d;
        for (int i = 0; i < rowSize - 1; i++)
        {
            Map<String, String> v = data.get(i);
            Double vc = Double.parseDouble(v.get(columnName));
            if(i%1000==0)
            {
                logger.info("sort:"+i+",total:"+(rowSize-1)+",per: "+(int)(i/(1.0*(rowSize-1))*10000)/100.0+"%,耗时："+(System.currentTimeMillis() - t) / 1000.0 + "秒");
            }
            if (vc < baseValue)
            {
                int j = baseIndex;
                while (j > 0)
                {
                    d = Double.parseDouble(sdata.get(j).get(columnName));
                    if (vc >= d)
                        break;
                    j--;
                }
                sdata.add(j, v);
                baseIndex++;
            }
            else
            {
                int j = baseIndex;
                while (j < sdata.size())
                {
                    d = Double.parseDouble(sdata.get(j).get(columnName));
                    if (vc < d)
                        break;
                    j++;
                }
                sdata.add(j, v);
            }
        }
        logger.info("排序列：" + columnName + " ,耗时：" + (System.currentTimeMillis() - t) / 1000.0 + "秒");
    }

//    private static void insertData(LinkedList<Map<String, String>> linkedList, in)

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

    private String getDateCellValue(HSSFCell cell)
    {
        String result = "";
        try
        {
            int cellType = cell.getCellType();
            if (cellType == HSSFCell.CELL_TYPE_NUMERIC)
            {
                Date date = cell.getDateCellValue();
                result = (date.getYear() + 1900) + "-" + (date.getMonth() + 1)
                        + "-" + date.getDate();
            }
            else if (cellType == HSSFCell.CELL_TYPE_STRING)
            {
                String date = getStringCellValue(cell);
                result = date.replaceAll("[年月]", "-").replace("日", "").trim();
            }
            else if (cellType == HSSFCell.CELL_TYPE_BLANK)
            {
                result = "";
            }
        }
        catch (Exception e)
        {
            System.out.println("日期格式不正确!");
            e.printStackTrace();
        }
        return result;
    }

    private String getCellFormatValue(HSSFCell cell)
    {
        String cellvalue = "";
        if (cell != null)
        {
            // 判断当前Cell的Type
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
}
