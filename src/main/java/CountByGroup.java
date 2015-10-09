import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by dev on 15-10-9.
 */
public class CountByGroup
{
    public static void main(String[] args) throws Exception
    {
        String path = args[0];
        Path filePath = Paths.get(args[0]);
        Map<String, Integer> countMapResult = new HashMap<>();
        HandleData handleData = new HandleData();
        handleData.path = filePath;
        String t0 = args[0].substring(args[0].lastIndexOf("."));
        handleData.resultPathStr = filePath.toRealPath().toString().replace(t0, "_result.xlsx");
        InputStream inputStream = new FileInputStream(filePath.toRealPath().toString());
        String[] title = handleData.title = handleData.readExcelTitle(inputStream);
        handleData.data1 = new ArrayList<>();
        handleData.readExcelContent(handleData.data1);
        int size = handleData.data1.size();
        handleData.data2 = new ArrayList<>();
        int fixedsize = 0;
        Map<String, Object> item2 = new HashMap<>();
        String id1 = title[0];
        String id2 = title[3];
        String id3 = title[title.length-1];
        for (int i = 0; i < size; i++)
        {
            Map<String, Object> item = handleData.data1.get(i);
            String id = String.valueOf(item.get(id1)) + String.valueOf(item.get(id2));
            if (!countMapResult.containsKey(id))
            {
                countMapResult.put(id, 1);
            }
            else
            {
                countMapResult.put(id, countMapResult.get(id) + 1);
            }
        }
        for (int i = 0; i < size; i++)
        {
            Map<String, Object> item = handleData.data1.get(i);
            String id = String.valueOf(item.get(id1)) + String.valueOf(item.get(id2));
            item.put(id3,String.valueOf(countMapResult.get(id)));
        }
        handleData.writeResultData(handleData.data1);
    }
}
