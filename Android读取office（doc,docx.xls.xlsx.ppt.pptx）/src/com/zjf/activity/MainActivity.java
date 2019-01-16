package com.zjf.activity;

import java.io.File;
import java.util.ArrayList;

import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.Menu;
import android.view.View;
import android.view.Window;
import android.widget.AdapterView;
import android.widget.AdapterView.OnItemClickListener;
import android.widget.ListView;

import com.zjf.adapter.WordAdapter;

public class MainActivity extends Activity implements OnItemClickListener {

	public static String sdPath = Environment.getExternalStorageDirectory()
			.getAbsolutePath();
	private ListView wordList;
	private WordAdapter wordAdapter;
	private ArrayList<File> fileList;
	
	private Intent intent;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		this.requestWindowFeature(Window.FEATURE_NO_TITLE);
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		wordList = (ListView) this.findViewById(R.id.list);

		fileList = new ArrayList<File>();
		getFiles(sdPath);
		for (int i = 0; i < fileList.size(); i++) {
			Log.e("文件名：",
					fileList.get(i).getName() + "\n文件路径："
							+ fileList.get(i).getAbsolutePath());
		}
		wordAdapter = new WordAdapter(this, fileList);
		wordList.setAdapter(wordAdapter);
		wordList.setOnItemClickListener(this);

	}

	// @Override
	// protected void onStart() {
	// fileList = new ArrayList<File>();
	// getFiles(sdPath);
	// for (int i = 0; i < fileList.size(); i++) {
	// Log.e("文件名：",
	// fileList.get(i).getName() + "\n文件路径："
	// + fileList.get(i).getAbsolutePath());
	// }
	// wordAdapter = new WordAdapter(this, fileList);
	// wordList.setAdapter(wordAdapter);
	// wordList.setOnItemClickListener(this);
	// super.onStart();
	// }

	public void getFiles(String path) {
		File file = new File(path);
		File[] files = file.listFiles();
		for (int i = 0; i < files.length; i++) {
			// Log.i("=====", files[i].getName() + "");
			if (files[i].isDirectory()) {
				getFiles(files[i].getAbsolutePath());
			} else {
				if (isWord(files[i])) {
					fileList.add(files[i]);
				}
			}
		}
	}

	public boolean isWord(File file) {
		String fileName = file.getName();
		if (fileName.endsWith(".doc") || fileName.endsWith(".docx")
				|| fileName.endsWith(".xls") || fileName.endsWith(".xlsx")
				|| fileName.endsWith(".pptx") || fileName.endsWith(".ppt")) {
			return true;
		}
		return false;
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}

	
	@Override
	public void onItemClick(AdapterView<?> arg0, View arg1, int arg2, long arg3) {
		
		String aa = fileList.get(arg2).getAbsolutePath();
		if (aa.endsWith(".pptx")||aa.endsWith(".ppt") ) {
			intent = new Intent(MainActivity.this, PPTActivity.class);
			intent.putExtra("filePath", fileList.get(arg2).getAbsolutePath());
			startActivity(intent);
		}if (aa.endsWith(".ppt")  ) {
//			intent = new Intent(MainActivity.this, PPTMainActivity.class);
//			intent.putExtra("filePath", fileList.get(arg2).getAbsolutePath());
//			startActivity(intent);
		}  
		else {
			intent = new Intent(MainActivity.this, ShowWordActivity.class);
			intent.putExtra("filePath", fileList.get(arg2).getAbsolutePath());
			startActivity(intent);
		}
		
		
//		Intent intent = new Intent();
//		intent.setClass(MainActivity.this, ShowWordActivity.class);
//		intent.putExtra("filePath", fileList.get(arg2).getAbsolutePath());
//		startActivity(intent);
	}

}
