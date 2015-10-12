import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.Iterator;

public class ExcelFileConverter
{


    private void convert(Workbook wb, FileOutputStream fos)
    {
        System.out.println("Excel--格式化--->Text");
        int count = wb.getNumberOfSheets();
        Sheet sheet;
        StringBuilder lineContent = new StringBuilder("");
        StringBuilder sep_t = new StringBuilder("\t");
        StringBuilder sep_n = new StringBuilder("\n");
        //int k = 50;
        for (int i = 0; i < count; i++)
        {
            sheet = wb.getSheetAt(i);
            float totalRow = Float.valueOf(sheet.getLastRowNum());
            int j = 0;
            try
            {
                for (Iterator<Row> iter = sheet.rowIterator(); iter.hasNext(); )
                {
                    Row row = iter.next();
                    for (Iterator<Cell> iter2 = row.cellIterator(); iter2.hasNext(); )
                    {
                        Cell cell = iter2.next();
                        Object value;
                        if (cell.getCellType() == Cell.CELL_TYPE_STRING)
                        {
                            value = cell.getStringCellValue();
                        }
                        else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
                        {
                            DecimalFormat df = new DecimalFormat("#");
                            value = df.format(cell.getNumericCellValue());
                        }
                        else
                        {
                            value = "";
                        }
                        lineContent.append(value).append(sep_t);
                    }
                    lineContent.append(sep_n);
                    if (j % 500 == 0)
                    {
                        System.out.println("sheet:" + (i + 1) + "/" + count + " progress:" + (j + 1) + "/" + totalRow + " " + ((int) ((j + 1) / totalRow * 10000)) / 100.0 + "%");
                    }
                    j++;
                }
                fos.write(lineContent.toString().getBytes());
                lineContent.delete(0, lineContent.length());
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }


}

