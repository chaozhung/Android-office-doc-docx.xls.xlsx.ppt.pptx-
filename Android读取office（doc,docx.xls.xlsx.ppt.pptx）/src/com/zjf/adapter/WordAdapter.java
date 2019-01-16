package com.zjf.adapter;

import java.io.File;
import java.util.ArrayList;

import com.zjf.activity.R;

import android.content.Context;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.BaseAdapter;
import android.widget.ImageView;
import android.widget.TextView;

public class WordAdapter extends BaseAdapter {

    private ArrayList<File> fileList;
    private Context context;

    public WordAdapter(Context context, ArrayList<File> fileList) {
        this.fileList = fileList;
        this.context = context;
    }

    @Override
    public int getCount() {
        return fileList.size();
    }

    @Override
    public Object getItem(int arg0) {
        return fileList.get(arg0);
    }

    @Override
    public long getItemId(int arg0) {
        return arg0;
    }

    @Override
    public View getView(int arg0, View arg1, ViewGroup arg2) {
        if (arg1 == null) {
            arg1 = LayoutInflater.from(context).inflate(R.layout.word_item,
                    null);
        }
        ImageView wordImage = (ImageView) arg1.findViewById(R.id.picture);
        TextView wordName = (TextView) arg1.findViewById(R.id.text);

        wordName.setText(fileList.get(arg0).getName());
        String aa = wordName.getText().toString();
        if (aa.endsWith(".doc") || aa.endsWith(".docx")) {
            wordImage.setImageDrawable(context.getResources().getDrawable(
                    R.drawable.word));
        } else if (aa.endsWith(".xls") || aa.endsWith(".xlsx")) {
            wordImage.setImageDrawable(context.getResources().getDrawable(
                    R.drawable.excel));
        } else if (aa.endsWith(".ppt") || aa.endsWith(".pptx")) {
            wordImage.setImageDrawable(context.getResources().getDrawable(
                    R.drawable.ppt));

        }

        return arg1;
    }

}
