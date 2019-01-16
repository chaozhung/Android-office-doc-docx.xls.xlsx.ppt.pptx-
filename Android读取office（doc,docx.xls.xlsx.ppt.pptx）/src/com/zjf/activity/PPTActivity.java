package com.zjf.activity;

import android.app.Activity;
import android.os.Bundle;

import com.itsrts.pptviewer.PPTViewer;

public class PPTActivity extends Activity {
	PPTViewer pptViewer;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_ppt);
		pptViewer = (PPTViewer) findViewById(R.id.pptviewer);
		String filePath = this.getIntent().getExtras().getString("filePath");
		pptViewer.setNext_img(R.drawable.next).setPrev_img(R.drawable.prev)
				.setSettings_img(R.drawable.settings)
				.setZoomin_img(R.drawable.zoomin)
				.setZoomout_img(R.drawable.zoomout);
		pptViewer.loadPPT(this, filePath);

	}

}
