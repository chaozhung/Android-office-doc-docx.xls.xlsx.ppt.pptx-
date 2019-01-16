package com.zjf.activity;

import java.io.BufferedInputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.apache.http.util.ByteArrayBuffer;
import org.apache.http.util.EncodingUtils;

import android.annotation.SuppressLint;
import android.app.Activity;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.KeyEvent;
import android.view.Window;
import android.webkit.WebSettings;
import android.webkit.WebSettings.LayoutAlgorithm;
import android.webkit.WebView;
import android.widget.Toast;

import com.zjf.util.WordToHtml;

@SuppressLint({ "SetJavaScriptEnabled", "NewApi" })
public class ShowWordActivity extends Activity {

	private WebView mOfficeWebview;
	private String mOfficeHtmlPath;
	public static String mSdcardPath = null;
	private String htmlPath = null;
	private String filePath = null;

	FR fr;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		this.requestWindowFeature(Window.FEATURE_NO_TITLE);
		super.onCreate(savedInstanceState);
		setContentView(R.layout.office_show_content);
		mOfficeWebview = (WebView) findViewById(R.id.webview);
		// 支持JavaScript
		mOfficeWebview.getSettings().setJavaScriptEnabled(true);
		// 设置可以支持缩放(双击以及拖动)
		mOfficeWebview.getSettings().setSupportZoom(true);
		mOfficeWebview.getSettings().setBuiltInZoomControls(true);
		// 不显示WebView缩放按钮(3.0以上版本)
		mOfficeWebview.getSettings().setDisplayZoomControls(false);
		// 扩大比例的缩放
		mOfficeWebview.getSettings().setUseWideViewPort(true);
		// 自适应屏幕
		mOfficeWebview.getSettings().setLayoutAlgorithm(
				LayoutAlgorithm.SINGLE_COLUMN);
		mOfficeWebview.getSettings().setLoadWithOverviewMode(true);
		WebSettings settings = mOfficeWebview.getSettings();
		settings.setLayoutAlgorithm(LayoutAlgorithm.SINGLE_COLUMN);
		mSdcardPath = Environment.getExternalStorageDirectory()
				.getAbsolutePath();
		filePath = this.getIntent().getExtras().getString("filePath");

		if (filePath.endsWith(".doc")) {
			Toast.makeText(ShowWordActivity.this, "doc", Toast.LENGTH_SHORT)
					.show();
			if (!isExistsDoc(filePath)) {
				try {
					Log.e("文件不存在", "HTML路径 ：" + mOfficeHtmlPath + "| "
							+ filePath);
					// String officeName = mOfficeHtmlPath.substring(
					// mOfficeHtmlPath.indexOf("0") + 1,
					// mOfficeHtmlPath.length());
					String[] mStrPath = mOfficeHtmlPath.split("/");
					String officeName = mStrPath[mStrPath.length - 1];
					htmlPath = mSdcardPath + File.separator + "gdemm"
							+ File.separator + officeName;
					WordToHtml.convert2Html(filePath, htmlPath);
				} catch (TransformerException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				} catch (ParserConfigurationException e) {
					e.printStackTrace();
				}
			}
			mOfficeWebview.loadUrl("file://" + htmlPath);
			// FR fr = new FR(filePath);
			// mOfficeWebview.loadUrl(fr.returnPath);
		}
		if (filePath.endsWith(".docx")) {
			Toast.makeText(ShowWordActivity.this, "docx", Toast.LENGTH_SHORT)
					.show();
			fr = new FR(filePath);
			mOfficeWebview.loadUrl(fr.returnPath);
			// readDocx();
		}
		if (filePath.endsWith(".xls")) {
			Toast.makeText(ShowWordActivity.this, "xls", Toast.LENGTH_SHORT)
					.show();
			fr = new FR(filePath);
			mOfficeWebview.loadUrl(fr.returnPath);
		}
		if (filePath.endsWith(".xlsx")) {
			Toast.makeText(ShowWordActivity.this, "xlsx", Toast.LENGTH_SHORT)
					.show();
			fr = new FR(filePath);
			mOfficeWebview.loadUrl(fr.returnPath);
		}
		if (filePath.endsWith(".pptx")) {
			Toast.makeText(ShowWordActivity.this, "pptx", Toast.LENGTH_SHORT)
					.show();
			fr = new FR(filePath);
			mOfficeWebview.loadUrl(fr.returnPath);
		}
		// Log.v("====================", htmlPath);
		// Log.e("HTMl文本内容：", getHtmlString("file://" + htmlPath));
		// String text = getHtmlString("file://" + htmlPath);
		// System.out.println(text);
		// String t = text.replace("vnf", "<a href = 'http://www.baidu.com'>" +
		// "vnf" + "</a>");
		// Log.e("HTMl文本内容：" ,"========================" + t);
		// File file = new File(htmlPath);
		// print(file,t);

	}

	public boolean isExistsDoc(String path) {
		mOfficeHtmlPath = path.replace(".doc", ".html");
		File file = new File(mOfficeHtmlPath);
		if (file.exists()) {
			return true;
		}
		return false;
	}

	public boolean isExistsDocx(String path) {
		mOfficeHtmlPath = path.replace(".docx", ".html");
		File file = new File(mOfficeHtmlPath);
		if (file.exists()) {
			return true;
		}
		return false;
	}

	// /**
	// * 读取docx文档
	// */
	// @SuppressWarnings("deprecation")
	// private void readDocx() {
	// try {
	// final LoadFromZipNG loader = new LoadFromZipNG();
	// WordprocessingMLPackage wordMLPackage = (WordprocessingMLPackage) loader
	// .get(new FileInputStream(filePath));
	// String IMAGE_DIR_NAME = "images";
	// String baseURL = this
	// .getDir(IMAGE_DIR_NAME, Context.MODE_WORLD_READABLE)
	// .toURL().toString();
	// System.out.println(baseURL);
	// ConversionImageHandler conversionImageHandler = new FileImageHandler(
	// IMAGE_DIR_NAME, baseURL, false, this);
	// HtmlExporterNonXSLT withoutXSLT = new HtmlExporterNonXSLT(
	// wordMLPackage, conversionImageHandler);
	// String html = XmlUtils.w3CDomNodeToString(withoutXSLT.export());
	// mOfficeWebview.loadDataWithBaseURL(baseURL, html, "text/html",
	// null, null);
	// } catch (Exception e) {
	// e.printStackTrace();
	// finish();
	// }
	// }

	/**
	 * 获取html中的内容
	 * 
	 * @param urlString
	 * @return
	 */
	public String getHtmlString(String urlString) {
		try {
			URL url = null;
			url = new URL(urlString);
			URLConnection ucon = null;
			ucon = url.openConnection();
			InputStream instr = null;
			instr = ucon.getInputStream();
			BufferedInputStream bis = new BufferedInputStream(instr);
			ByteArrayBuffer baf = new ByteArrayBuffer(500);
			int current = 0;
			while ((current = bis.read()) != -1) {
				baf.append((byte) current);
			}
			return EncodingUtils.getString(baf.toByteArray(), "UTF-8");
		} catch (Exception e) {
			return "";
		}

	}

	public void print(File file, String str) {
		FileWriter fw = null;
		BufferedWriter bw = null;
		try {
			fw = new FileWriter(file, false);
			bw = new BufferedWriter(fw);
			bw.write(str);
			bw.flush();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				bw.close();
				fw.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * 删除HTML文件
	 * 
	 * @param fileName
	 */
	public static void delFile(String fileName) {
		File file = new File(fileName);
		if (file.isFile()) {
			file.delete();
		}
		file.exists();
	}

	public void RecursionDeleteFile(File file) {
		if (file.isFile()) {
			file.delete();
			return;
		}
		if (file.isDirectory()) {
			File[] childFile = file.listFiles();
			if (childFile == null || childFile.length == 0) {
				file.delete();
				return;
			}
		}
	}

	@Override
	public boolean onKeyDown(int keyCode, KeyEvent event) {
		if (keyCode == KeyEvent.KEYCODE_BACK && event.getRepeatCount() == 0) { // 按下的如果是BACK，同时没有重复

			deleteDir();
			finish();
			return true;
		}

		return super.onKeyDown(keyCode, event);
	}

	public static void deleteDir() {
		String SDPATH = Environment.getExternalStorageDirectory()
				+ "/gdemm/aaa/";
		File dir = new File(SDPATH);
		if (dir == null || !dir.exists() || !dir.isDirectory())
			return;

		for (File file : dir.listFiles()) {
			if (file.isFile())
				file.delete(); // 删除所有文件
			else if (file.isDirectory())
				deleteDir(); // 递规的方式删除文件夹
		}
		dir.delete();// 删除目录本身
	}
}
