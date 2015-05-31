import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;

/**
 * Created by azure on 2015/5/30 0030.
 */
public class SortData
{
    public static void main(String[] args) throws Exception
    {
        Path path = Paths.get(args[0]);
        HandleData handleData = new HandleData();
        handleData.resultPathStr = path.toRealPath().toString().replace(".xls", "_result.xls");
        handleData.column1 = args[1];
        handleData.column2 = args[2];
        handleData.logger.info(path + "," + path.toFile().exists() + "," + handleData.column1 + "," + handleData.column2);
        InputStream inputStream = new FileInputStream(path.toRealPath().toString());
        handleData.title =  handleData.readExcelTitle(inputStream);
        handleData.readExcelContent(null);
        handleData.data1 = new ArrayList<>();
        handleData.sortData(handleData.data1, handleData.column1);
        handleData.sortData(handleData.data1, handleData.column2);
//        filterData(data1, column2);
//        //  printData(data1, column1);
//        data2 = new ArrayList<>();
//        sortData(data2, column2);
//        filterData(data2, column1);
        // printData(data2, column2);
        handleData.resultData = new HashMap<>();
        HandleData.logger.info("开始mergeData...");
        handleData.mergeData(handleData.resultData, handleData.data1);
        //mergeData(resultData, data2);
        HandleData.logger.info("完成mergeData...共计：" + handleData.resultData.size() + "行");
        handleData.writeResultData(null);
    }
}
