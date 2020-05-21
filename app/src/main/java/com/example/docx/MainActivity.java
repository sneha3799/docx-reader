package com.example.docx;

import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.os.Bundle;
import android.text.method.ScrollingMovementMethod;
import android.util.Log;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.widget.Button;
import android.widget.ImageView;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    Button btn,btn1;
    public static InputStream InputstreamGeneralPhotos1;
    public static InputStream InputstreamGeneralPhotos2;
    public static InputStream InputstreamGeneralPhotos3;
    public static InputStream InputstreamGeneralPhotos4;
    public static InputStream InputstreamGeneralPhotos5;
    public static InputStream InputstreamGeneralPhotos6;
    public static InputStream InputstreamBicyclePhotos1;
    public static InputStream InputstreamBicyclePhotos2;
    public static InputStream InputstreamFourhoursPhotos1;
    public static InputStream InputstreamFourhoursPhotos2;
    public static InputStream InputstreamFourhoursPhotos3;
    public static InputStream InputstreamFourhoursPhotos4;

    static {
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLInputFactory",
                "com.fasterxml.aalto.stax.InputFactoryImpl"
        );
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLOutputFactory",
                "com.fasterxml.aalto.stax.OutputFactoryImpl"
        );
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLEventFactory",
                "com.fasterxml.aalto.stax.EventFactoryImpl"
        );
    }

    TextView textView;
    ImageView img_generalPhotos1,img_generalPhotos2,img_generalPhotos3,img_generalPhotos4,img_generalPhotos5,img_generalPhotos6,img_generalPhotos7,img_generalPhotos8,img_generalPhotos9,img_generalPhotos10,img_generalPhotos11,img_generalPhotos12;

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        MenuInflater inflater = getMenuInflater();
        inflater.inflate(R.menu.menu,menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        switch(item.getItemId()){
            case R.id.item:
                openDocumentFromFileManager();
                return true;
//            case R.id.item1:
//                return true;
            default:
                return super.onOptionsItemSelected(item);
        }
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        textView = (TextView)findViewById(R.id.textView);

    }

    private void openDocumentFromFileManager() {
        //this is the action to open doc file from file manager
        Intent i = new Intent();
        i.setType("application/*");
        i.setAction(Intent.ACTION_GET_CONTENT);
        if (PermissionsHelper.getPermission(this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE, R.string.title_storage_permission
                , R.string.text_storage_permission, 1111)) {
            startActivityForResult(Intent.createChooser(i, "Select Document"), 111);
        }
    }


    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        try {
            if (resultCode==RESULT_OK){
                switch (requestCode){
                    case 111:
                        //this is action performed after openDocumentFromFileManager() when doc is selected
                        FileInputStream inputStream = (FileInputStream) getContentResolver().openInputStream(data.getData());
                        XWPFDocument docx = new XWPFDocument(inputStream);

                        extractImages(docx);
//                        extractText(docx);

                        List<XWPFParagraph> paragraphList = docx.getParagraphs();

                        for (XWPFParagraph paragraph:paragraphList){
                            String paragrapthText = paragraph.getText();
                            System.out.println(paragrapthText);
                        }


                }
            }else {
                Toast.makeText(this, "Your file is not loaded", Toast.LENGTH_SHORT).show();
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void extractImages(XWPFDocument docx){
        try {

            //function to retrieve all the images from the doc file that have been selected
//
//            HWPFDocument wordDoc = new HWPFDocument(new FileInputStream(fileName));
//
            XWPFWordExtractor extractor = new XWPFWordExtractor(docx);

//            extractText(extractor);

            System.out.println(extractor.getText());

            textView.setText(extractor.getText());
            textView.setMovementMethod(new ScrollingMovementMethod());

            int pages = docx.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();
            Log.i("Pages",Integer.toString(pages));

            List<XWPFPictureData> picList = docx.getAllPackagePictures();
            int s = picList.size();
            Log.i("Images",Integer.toString(s));

            int i = 1;

            for (XWPFPictureData pic : picList) {

                System.out.print(pic.getPictureType());
                System.out.print(pic.getData());
                System.out.println("Image Number: " + i + " " + pic.getFileName());

                switch (i){
                    case 1:
                        InputstreamGeneralPhotos1 = new ByteArrayInputStream(pic.getData());
                        if (InputstreamGeneralPhotos1 != null){
                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos1);
                            img_generalPhotos1.setImageBitmap(selectedImage);
//                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
                        }
                        break;
                    case 2:
                        InputstreamGeneralPhotos2 = new ByteArrayInputStream(pic.getData());
                        if (InputstreamGeneralPhotos2 != null){
                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos2);
                            img_generalPhotos2.setImageBitmap(selectedImage);
//                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
                        }
                        break;
                    case 3:
                        InputstreamGeneralPhotos3 = new ByteArrayInputStream(pic.getData());
                        if (InputstreamGeneralPhotos3 != null){
                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos3);
                            img_generalPhotos3.setImageBitmap(selectedImage);
//                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
                        }
                        break;
                    case 4:
                        InputstreamGeneralPhotos4 = new ByteArrayInputStream(pic.getData());
                        if (InputstreamGeneralPhotos4 != null){
                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos4);
                            img_generalPhotos4.setImageBitmap(selectedImage);
//                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
                        }
                        break;
                    case 5:
                        InputstreamGeneralPhotos5 = new ByteArrayInputStream(pic.getData());
                        if (InputstreamGeneralPhotos5 != null){
                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos5);
                            img_generalPhotos5.setImageBitmap(selectedImage);
//                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
                        }
                        break;
                    case 6:
                        InputstreamGeneralPhotos6 = new ByteArrayInputStream(pic.getData());
                        if (InputstreamGeneralPhotos6 != null){
                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos6);
                            img_generalPhotos6.setImageBitmap(selectedImage);
//                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
                        }
                        break;
                    case 7:
                        InputstreamBicyclePhotos1 = new ByteArrayInputStream(pic.getData());
                        break;
                    case 8:
                        InputstreamBicyclePhotos2 = new ByteArrayInputStream(pic.getData());
                        break;
                    case 9:
                        InputstreamFourhoursPhotos1 = new ByteArrayInputStream(pic.getData());
                        break;
                    case 10:
                        InputstreamFourhoursPhotos2 = new ByteArrayInputStream(pic.getData());
                        break;
                    case 11:
                        InputstreamFourhoursPhotos3 = new ByteArrayInputStream(pic.getData());
                        break;
                    case 12:
                        InputstreamFourhoursPhotos4 = new ByteArrayInputStream(pic.getData());
                        break;
                }
                i++;
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }
}
