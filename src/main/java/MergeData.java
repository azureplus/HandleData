import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;

/**
 * Created by azure on 2015/5/30 0030.
 */
public class MergeData
{
    /**
     * args[0]:Excel_1 path
     * args[1]:Excel_2 path
     * args[2]:0,1
     * args[3]:merge column
     *
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception
    {
        Path path1 = null;
        Path path2 = null;
        String path = MergeData.class.getClassLoader().getResource("").toString().replace("file:/", "");
        if (args[0].indexOf(":") == -1)
        {
            path1 = Paths.get(path + args[0]);
        } else
        {
            path1 = Paths.get(args[0]);
        }
        if (args[1].indexOf(":") == -1)
        {
            path2 = Paths.get(path + args[1]);
        } else
        {
            path2 = Paths.get(args[1]);
        }

        HandleData handleData = new HandleData();
        handleData.column1 = args[3];
        handleData.path = path1;
        handleData.resultPathStr = path1.toRealPath().toString().replace(".xls", "_result.xls");
        HandleData.logger.info(path1 + "," + path1.toFile().exists() + "," + handleData.column1 + "," + handleData.column2);
        InputStream inputStream = new FileInputStream(path1.toRealPath().toString());
        String[] title1 = handleData.title = handleData.readExcelTitle(inputStream);
        handleData.data1 = new ArrayList<>();
        handleData.readExcelContent(handleData.data1);

        handleData.path = path2;
        inputStream = new FileInputStream(path2.toRealPath().toString());
        String[] title2 = handleData.title = handleData.readExcelTitle(inputStream);
        handleData.data2 = new ArrayList<>();
        handleData.readExcelContent(handleData.data2);
        handleData.updateUID(handleData.data1, "year", "id");
        handleData.updateUID(handleData.data2, "year", "id");
        handleData.resultData = new HashMap<>();
        handleData.mergeData(handleData.resultData, handleData.data1);
        handleData.mergeData(handleData.resultData, handleData.data2, handleData.column1);
        String[] title = new String[title1.length + 1];
        title = Arrays.copyOf(title1, title.length);
        title[title.length-1] = handleData.column1;
        handleData.title = title;
        HandleData.logger.info(Arrays.toString(title));
        handleData.data2 = new ArrayList<>();
        handleData.sortData(handleData.resultData, handleData.data2 ,handleData.title[0]);
        handleData.writeResultData(handleData.data2);
    }
}
