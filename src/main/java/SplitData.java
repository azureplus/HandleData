import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

/**
 * Created by azure on 15/9/26.
 */
public class SplitData
{
    public static void main(String[] args) throws Exception
    {
        Path filePath = Paths.get(args[0]);

        HandleData handleData = new HandleData();
        handleData.path = filePath;
        String t0  = args[0].substring(args[0].lastIndexOf("."));
        handleData.resultPathStr = filePath.toRealPath().toString().replace(t0, "_result.xlsx");
        InputStream inputStream = new FileInputStream(filePath.toRealPath().toString());
        String[] title = handleData.title = handleData.readExcelTitle(inputStream);
        handleData.data1 = new ArrayList<>();
        handleData.readExcelContent(handleData.data1);
        int size = handleData.data1.size();
        handleData.data2 = new ArrayList<>();
        int fixedsize = 0;
        Map<String, Object> item2 = new HashMap<>();
        int i = 0;
        int s, s1;
        int p1 = 9;
        int p2 = 10;
        try
        {
            for (i = 0; i < size; i++)
            {
                Map<String, Object> item = handleData.data1.get(i);
                if (i < 2)
                {
                    handleData.data2.add(item);
                    continue;
                }
                String date = String.valueOf(item.get(title[p1]));
                String d7 = String.valueOf( item.get(title[p2]));
                if(date.indexOf("N/A") != -1 || d7 .indexOf("N/A") != -1)
                {
                    item.put(title[0], "fixed" + item.get(title[0]));

                    handleData.data2.add(item);
                    continue;
                }
                String[] d = date.split(",");
                if (d.length > 1)
                {
                    item.put(title[0], "fixed" + item.get(title[0]));
                    item.put(title[p1], d[0]);
                    item.put(title[p2], d[1]);
                    fixedsize++;
                } else if (d7 == null || d7 == "")
                {
                    item.put(title[0], "fixed" + item.get(title[0]));
                    handleData.data2.add(item);
                    continue;
                }
                String t = String.valueOf(item.get(title[p1])).split("-")[0];
                s = (int)(double)Double.valueOf(t);
                t = String.valueOf(item.get(title[p2])).split("-")[0];
                s1 = (int)(double)Double.valueOf(t);
                item.put(title[p1], String.valueOf(s));
                item.put(title[p2], String.valueOf(s1));
                handleData.data2.add(item);
                for (int j = s; j < s1; j++)
                {
                    item2 = new HashMap<>(title.length);
                    item2.putAll(item);
                    item2.put(title[p1], String.valueOf(j+1));
                    handleData.data2.add(item2);
                }

            }
        } catch (Exception e)
        {
            System.out.println(i);
            e.printStackTrace();
        }
        handleData.data1 = null;
        handleData.writeResultData(handleData.data2);
        System.out.printf("========= fixedsize: " + fixedsize);
    }
}
