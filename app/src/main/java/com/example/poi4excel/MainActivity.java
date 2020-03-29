package com.example.poi4excel;

import android.annotation.SuppressLint;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.widget.TextView;

import androidx.appcompat.app.AppCompatActivity;

import com.yanzhenjie.permission.AndPermission;
import com.yanzhenjie.permission.runtime.Permission;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

public class MainActivity extends AppCompatActivity {
    private static final String TAG = "poi4excel";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
    }

    @SuppressLint("WrongConstant")
    @Override
    protected void onResume() {
        super.onResume();

        AndPermission.with(this)
                .runtime()
                .permission(Permission.Group.STORAGE)
                .onGranted(permissions -> {
                    readFiles();
                })
                .onDenied(permissions -> {
                    finish();
                })
                .start();
    }

    private void readFiles() {
        boolean isSDCardExist = Environment.getExternalStorageState().equals(android.os.Environment.MEDIA_MOUNTED);
        if (!isSDCardExist) {
            Log.e(TAG, "SDCard is not exist");
            return;
        }

        copyAssetsToDst("", "");

        File sdDir = Environment.getExternalStorageDirectory();
        String filePath = sdDir.getAbsolutePath() + File.separator + "PoiTest.xls";
//        String filePath = sdDir.getAbsolutePath() + File.separator + "PoiTest.xlsx";
        File file = new File(filePath);
        if (!file.exists()) {
            Log.e(TAG, "Excel file is exist: " + file.exists());
            return;
        }

        try {
            Workbook workbook;
            String extString = filePath.substring(filePath.lastIndexOf("."));
            if (".xls".equals(extString)) {
                workbook = new HSSFWorkbook(new FileInputStream(filePath));
            } else if (".xlsx".equals(extString)) {
                workbook = new XSSFWorkbook(new FileInputStream(filePath));
            } else {
                workbook = null;
            }

            if (workbook != null) {
                String tmp = "";
                for (int j = 0; j < workbook.getNumberOfSheets(); j++) {
                    tmp += workbook.getSheetAt(j).getSheetName() + "  ";
                }
                Log.e(TAG, "All Sheet:  " + tmp);

                TextView showTv = findViewById(R.id.tv_show);
                showTv.setText(tmp);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void copyAssetsToDst(String srcPath, String dstPath) {
        try {
            String fileNames[] = getAssets().list(srcPath);
            if (fileNames.length > 0) {
                File file = new File(Environment.getExternalStorageDirectory(), dstPath);
                if (!file.exists()) {
                    file.mkdirs();
                }

                for (String fileName : fileNames) {
                    if (!srcPath.equals("")) { // assets 文件夹下的目录
                        copyAssetsToDst(srcPath + File.separator + fileName, dstPath + File.separator + fileName);
                    } else { // assets 文件夹
                        copyAssetsToDst(fileName, dstPath + File.separator + fileName);
                    }
                }
            } else {
                File outFile = new File(Environment.getExternalStorageDirectory(), dstPath);
                InputStream is = getAssets().open(srcPath);
                FileOutputStream fos = new FileOutputStream(outFile);
                byte[] buffer = new byte[1024];
                int byteCount;
                while ((byteCount = is.read(buffer)) != -1) {
                    fos.write(buffer, 0, byteCount);
                }
                fos.flush();
                is.close();
                fos.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
