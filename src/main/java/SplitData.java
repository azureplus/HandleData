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
public class SplitData {
    public static void main(String[] args) throws Exception{
        Path filePath = Paths.get(args[0]);

        HandleData handleData = new HandleData();
        handleData.path = filePath;
        handleData.resultPathStr = filePath.toRealPath().toString().replace(".xls", "_result.xlsx");
        InputStream inputStream = new FileInputStream(filePath.toRealPath().toString());
        String[] title = handleData.title = handleData.readExcelTitle(inputStream);
        handleData.data1 = new ArrayList<>();
        handleData.readExcelContent(handleData.data1);
        int size = handleData.data1.size();
        handleData.data2 = new ArrayList<>();
        int fixedsize = 0;
        Map<String,Object> item2 = new HashMap<>();
        int i=0;
        int s,s1;
        try
        {
            for ( i = 0; i < size; i++) {
                Map<String,Object> item = handleData.data1.get(i);
                String i6 = (String)item.get(title[6]);
                String i7 = (String)item.get(title[7]);
                if(i<3 || i6.indexOf("N/A")!=-1 ||i7.indexOf("N/A")!=-1)
                {
                    handleData.data2.add(item);
                    continue;
                }
                String date = (String)item.get(title[6]);
                String[] d = date.split(",");
                if(d.length>1)
                {
                    item.put(title[0],"fixed"+item.get(title[0]));
                    item.put(title[6],d[0]);
                    item.put(title[7],d[1]);
                    fixedsize++;
                }
                else if(item2.get(title[7]) ==null || item.get(title[7]) == "")
                {
                    item.put(title[0],"fixed"+item.get(title[0]));
                    handleData.data2.add(item);
                    continue;
                }
                String t = ((String) item.get(title[6])).split("-")[0];
                s = Integer.valueOf(t);
                s1 = Integer.valueOf(((String) item.get(title[7])).split("-")[0]);
                item.put(title[6],String.valueOf(s));
                item.put(title[7],String.valueOf(s1));
                handleData.data2.add(item);
                for (int j = s; j <= s1; j++) {
                    item2 = new HashMap<>(title.length);
                    item2.putAll(item);
                    item2.put(title[6],String.valueOf(j));
                    handleData.data2.add(item2);
                }

            }
        }
        catch (Exception e)
        {
            System.out.println(i);
            e.printStackTrace();
        }
        handleData.writeResultData(handleData.data2);
        System.out.printf("========="+fixedsize);
    }
}
