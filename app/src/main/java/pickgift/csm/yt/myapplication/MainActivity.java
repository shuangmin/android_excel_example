package pickgift.csm.yt.myapplication;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import android.content.Context;
import android.content.Intent;
import android.os.Build;
import android.os.Bundle;
import android.util.Log;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    String[] title = new String[]{"标题","标题","标题","标题","标题","标题","标题","标题","标题","标题"};
    String[] str = new String[]{"1","2","3","4","5","6","7","8","9","10"};
    String filename = "/sdcard/textExcel.xls";

    Integer[][] arr = new Integer[3][2];// {{1,2},{4,5},{7,8}};

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);


        Log.e(TAG, "onCreate: " + arr.length);

        int b = 0;
        for(int a=0;a<arr.length;a++){//控制每个一维数组
            for(int i=0;i<arr[a].length;i++){//控制每个一维数组中的元素
                arr[a][i] = b++;
            }
            System.out.println();//每执行完一个一维数组换行
        }

        for(int a=0;a<arr.length;a++){//控制每个一维数组
            for(int i=0;i<arr[a].length;i++){//控制每个一维数组中的元素
                System.out.print("onCreate " + arr[a][i]+" ");//输出每个元素的值
            }
            System.out.println();//每执行完一个一维数组换行
        }

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            requestPermissions(new String[]{"android.permission.WRITE_EXTERNAL_STORAGE"},0x1);
        }
        //exportExcelFile("/sdcard/myexcel.xls");
    }

    public double sum(Sheet sh, int cloumnIndex, int rowPosStart, int size) {
        double sum = 0.0;
        int rowPosEnd = rowPosStart + size;
        for(int row = rowPosStart; row < rowPosEnd; row++) {
            sum += Double.valueOf(sh.getRow(row).getCell(cloumnIndex).getStringCellValue());
            sh.getRow(row).getCell(row).getCellType();
        }
        return sum;
    }


    public double sumPrev(Sheet sh, int cloumnIndex, int endRowPos) {

        int endPos = endRowPos;

        double sum = 0.0;
        int effectiveSize = 0;
        while (endPos-- > 0) {
            try {
                double val = Double.valueOf(sh.getRow(endPos).getCell(cloumnIndex).getStringCellValue());
                if(val > 0.0 ) {
                    sum += val;
                    ++effectiveSize;
                }
            }catch (Throwable e) {
                break;
            }
        }

        Log.e(TAG, "sumPrev: cloumnIndex " + cloumnIndex + " endRowPos " + endRowPos + " sum " + sum);
        return effectiveSize > 0 ? sum / effectiveSize : 0;
    }

    private static final String TAG = "MainActivity";
    public void exportExcelFile(String fullPath){
        int size = 10;
        Workbook wb = new HSSFWorkbook();
        Sheet sh = wb.createSheet();

        Row row1 = sh.createRow(0);
        for(int cellnum=0;cellnum<10;cellnum++){
            Cell cell = row1.createCell(cellnum);
            cell.setCellValue(title[cellnum]);
        }

        for(int rownum=1;rownum<size + 1;rownum++){
            Row row = sh.createRow(rownum);
            for(int cellnum=0;cellnum<10;cellnum++){
                Cell cell = row.createCell(cellnum);
                cell.setCellValue(str[cellnum]);
            }
        }

//        Log.e(TAG, "exportExcelFile: 1  "  + Double.valueOf(sh.getRow(0).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 2  "  + Double.valueOf(sh.getRow(1).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 3  "  + Double.valueOf(sh.getRow(2).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 4  "  + Double.valueOf(sh.getRow(3).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 5  "  + Double.valueOf(sh.getRow(4).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 6  "  + Double.valueOf(sh.getRow(5).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 7  "  + Double.valueOf(sh.getRow(6).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 8  "  + Double.valueOf(sh.getRow(7).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 9  "  + Double.valueOf(sh.getRow(8).getCell(0).getStringCellValue()));
//        Log.e(TAG, "exportExcelFile: 10 " + Double.valueOf(sh.getRow(8).getCell(0).getStringCellValue()));

        //Log.e(TAG, "exportExcelFile: ====>>>>0  " + sum(sh,0,0,9));
        sumPrev(sh,0,0);
        sumPrev(sh,0,1);
        sumPrev(sh,0,2);
        sumPrev(sh,0,3);
        sumPrev(sh,0,4);
        sumPrev(sh,0,5);
        sumPrev(sh,0,6);
        sumPrev(sh,0,7);
        sumPrev(sh,0,8);
        sumPrev(sh,0,9);
        sumPrev(sh,0,10);
        sumPrev(sh,0,11);
        sumPrev(sh,0,12);
        sumPrev(sh,0,13);


        try {
            final String excelPath = fullPath + ".xls";
            FileOutputStream fos = new FileOutputStream(excelPath);//openFileOutput(filename, Context.MODE_PRIVATE);
            wb.write(fos);
            fos.close();
            Toast.makeText(getApplicationContext(),"导出成功",Toast.LENGTH_SHORT).show();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);

        exportExcelFile("/sdcard/00.xls");
    }
}
