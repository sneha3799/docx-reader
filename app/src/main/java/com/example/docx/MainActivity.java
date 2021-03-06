package com.example.docx;

import androidx.appcompat.app.AppCompatActivity;

import android.annotation.SuppressLint;
import android.content.ContentResolver;
import android.content.Context;
import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Color;
import android.graphics.Matrix;
import android.graphics.Typeface;
import android.media.Image;
import android.net.Uri;
import android.nfc.Tag;
import android.os.Bundle;
import android.os.Handler;
import android.provider.MediaStore;
import android.text.Layout;
import android.text.SpannableString;
import android.text.method.ScrollingMovementMethod;
import android.text.style.UnderlineSpan;
import android.util.DisplayMetrics;
import android.util.Log;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.view.View;
import android.webkit.MimeTypeMap;
import android.widget.Button;
import android.widget.ImageView;
import android.widget.LinearLayout;
import android.widget.RelativeLayout;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ICell;
import org.apache.poi.xwpf.usermodel.IRunElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFieldRun;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.XWPFSDTCell;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

//public class MainActivity extends AppCompatActivity {
//    LinearLayout mainUI;
//
//    Button btn,btn1;
//    public static InputStream InputstreamGeneralPhotos1;
//    public static InputStream InputstreamGeneralPhotos2;
//    public static InputStream InputstreamGeneralPhotos3;
//    public static InputStream InputstreamGeneralPhotos4;
//    public static InputStream InputstreamGeneralPhotos5;
//    public static InputStream InputstreamGeneralPhotos6;
//    public static InputStream InputstreamBicyclePhotos1;
//    public static InputStream InputstreamBicyclePhotos2;
//    public static InputStream InputstreamFourhoursPhotos1;
//    public static InputStream InputstreamFourhoursPhotos2;
//    public static InputStream InputstreamFourhoursPhotos3;
//    public static InputStream InputstreamFourhoursPhotos4;
//    List<XWPFPictureData> picList;
//    static {
//        System.setProperty(
//                "org.apache.poi.javax.xml.stream.XMLInputFactory",
//                "com.fasterxml.aalto.stax.InputFactoryImpl"
//        );
//        System.setProperty(
//                "org.apache.poi.javax.xml.stream.XMLOutputFactory",
//                "com.fasterxml.aalto.stax.OutputFactoryImpl"
//        );
//        System.setProperty(
//                "org.apache.poi.javax.xml.stream.XMLEventFactory",
//                "com.fasterxml.aalto.stax.EventFactoryImpl"
//        );
//    }
//
//    TextView textView;
//    ImageView img_generalPhotos1,img_generalPhotos2,img_generalPhotos3,img_generalPhotos4,img_generalPhotos5,img_generalPhotos6,img_generalPhotos7,img_generalPhotos8,img_generalPhotos9,img_generalPhotos10,img_generalPhotos11,img_generalPhotos12;
//
//    @Override
//    public boolean onCreateOptionsMenu(Menu menu) {
//        MenuInflater inflater = getMenuInflater();
//        inflater.inflate(R.menu.menu,menu);
//        return true;
//    }
//
//    @Override
//    public boolean onOptionsItemSelected(MenuItem item) {
//        switch(item.getItemId()){
//            case R.id.item:
//                openDocumentFromFileManager();
//                return true;
////            case R.id.item1:
////                return true;
//            default:
//                return super.onOptionsItemSelected(item);
//        }
//    }
//
//    @Override
//    protected void onCreate(Bundle savedInstanceState) {
//        super.onCreate(savedInstanceState);
//        setContentView(R.layout.activity_main);
//
//        mainUI = findViewById(R.id.MainUI);
//
////        textView = (TextView)findViewById(R.id.textView);
////        img_generalPhotos1 = (ImageView)findViewById(R.id.imageView);
////        img_generalPhotos2 = (ImageView)findViewById(R.id.imageView2);
//
//    }
//
//    private void openDocumentFromFileManager() {
//        //this is the action to open doc file from file manager
//        Intent i = new Intent();
//        i.setType("application/*");
//        i.setAction(Intent.ACTION_GET_CONTENT);
//        if (PermissionsHelper.getPermission(this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE, R.string.title_storage_permission
//                , R.string.text_storage_permission, 1111)) {
//            startActivityForResult(Intent.createChooser(i, "Select Document"), 111);
//        }
//    }
//
//
//    @Override
//    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
//        super.onActivityResult(requestCode, resultCode, data);
//
//        try {
//            if (resultCode == RESULT_OK) {
//                switch (requestCode) {
//                    case 111:
//                        //this is action performed after openDocumentFromFileManager() when doc is selected
//                        FileInputStream inputStream = (FileInputStream) getContentResolver().openInputStream(data.getData());
//
//                        Uri uri = data.getData();
//                        String extension = getMimeType(this, uri);
//                        Log.i("Extension", extension);
//
//                        XWPFDocument docx = new XWPFDocument(inputStream);
//                        traverseBodyElements(docx.getBodyElements());
//                        picList = docx.getAllPackagePictures();
//
//                        docx.close();
//
//
//                        List<XWPFParagraph> paragraphList = docx.getParagraphs();
//                        for (XWPFParagraph paragraph : paragraphList) {
//                            String paragrapthText = paragraph.getText();
//
//                            if(paragrapthText.length()>1) {
////                                addTextViews(paragrapthText);
////                                addElementsUI(paragrapthText,null);
//
//                            }
//                        }
//                }
//            }else {
//                Toast.makeText(this, "Your file is not loaded", Toast.LENGTH_SHORT).show();
//            }
//
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//
//    public void traversePictures(List<XWPFPicture> pictures)  {
//        for (XWPFPicture picture : pictures) {
//
//            //Create Image view to display images
//
////            Log.i("Desc",picture.getDescription());
//            System.out.println("Picture "+picture);
//            XWPFPictureData pictureData = picture.getPictureData();
//            Log.i("PictureData ", pictureData.toString());
//
//            long w = picture.getCTPicture().getSpPr().getXfrm().getExt().getCx();
//            long h = picture.getCTPicture().getSpPr().getXfrm().getExt().getCy();
//
//            addElements(pictureData,w,h);
////            addElementsUI(null,pictureData);
//        }
//
//    }
//
//    StringBuilder para = new StringBuilder();
//    private static ArrayList<StringBuilder> paras = new ArrayList<>();
//    private static int paraIndex = 0;
//
//    public void traverseRunElements(List<IRunElement> runElements) throws Exception {
//
//        System.out.println("PARAINDICES"+ paras.size());
//        System.out.println("TRAVERSE RUN ELEMENTS");
//
//        for (IRunElement runElement : runElements) {
//            if (runElement instanceof XWPFFieldRun) {
//                XWPFFieldRun fieldRun = (XWPFFieldRun)runElement;
//                System.out.println("fieldRunClassName "+fieldRun.getClass().getName());
//                System.out.println("fieldName "+fieldRun);
//                traversePictures(fieldRun.getEmbeddedPictures());
//            } else if (runElement instanceof XWPFHyperlinkRun) {
//                XWPFHyperlinkRun hyperlinkRun = (XWPFHyperlinkRun)runElement;
//                System.out.println("hyperLinkRunClassName "+ hyperlinkRun.getClass().getName());
//                System.out.println("hyperlinkRun "+hyperlinkRun);
//                traversePictures(hyperlinkRun.getEmbeddedPictures());
//            } else if (runElement instanceof XWPFRun) {
//                XWPFRun run = (XWPFRun)runElement;
//                System.out.println("runClassName "+run.getClass().getName());
//                System.out.println("run "+run);
//
//                //Appending text to paragraph
////                paras.get(paraIndex).append(run);
//                para.append(run);
//                paras.add(para);
////                System.out.println("PARA1TEXT "+para.toString());
////                addTextViews(run.toString(),run.getFontSize(),run.isBold(),run.getUnderline());
//
//                traversePictures(run.getEmbeddedPictures());
//
//            } else if (runElement instanceof XWPFSDT) {
//                XWPFSDT sDT = (XWPFSDT)runElement;
//                System.out.println("sDT"+sDT);
//                System.out.println("SDT_CONTENT "+sDT.getContent());
//                //ToDo: The SDT may have traversable content too.
//            }
//        }
//    }
//
//    public void traverseTableCells(List<ICell> tableICells) throws Exception {
//        for (ICell tableICell : tableICells) {
//            if (tableICell instanceof XWPFSDTCell) {
//                XWPFSDTCell sDTCell = (XWPFSDTCell)tableICell;
//                System.out.println("sDTCELL "+sDTCell);
//                //ToDo: The SDTCell may have traversable content too.
//            } else if (tableICell instanceof XWPFTableCell) {
//                XWPFTableCell tableCell = (XWPFTableCell)tableICell;
//                System.out.println("TableCell "+tableCell);
//                traverseBodyElements(tableCell.getBodyElements());
//            }
//        }
//    }
//
//    public void traverseTableRows(List<XWPFTableRow> tableRows) throws Exception {
//        for (XWPFTableRow tableRow : tableRows) {
//            System.out.println("TableRow "+tableRow);
//            traverseTableCells(tableRow.getTableICells());
//        }
//    }
//
//    public void traverseBodyElements(List<IBodyElement> bodyElements) throws Exception {
//        System.out.println("TRAVERSE BODY ELEMENTS");
//
//        for (IBodyElement bodyElement : bodyElements) {
//            if (bodyElement instanceof XWPFParagraph) {
//                XWPFParagraph paragraph = (XWPFParagraph)bodyElement;
//                System.out.println("PARA "+paragraph);
//
//                //Creating textView & paragraph using String Builder
//                paras.add(new StringBuilder());
//                traverseRunElements(paragraph.getIRuns());
//                paraIndex = paraIndex + 1;
//
//            } else if (bodyElement instanceof XWPFSDT) {
//                XWPFSDT sDT = (XWPFSDT)bodyElement;
//                System.out.println("SDT"+sDT);
//                System.out.println("SDT_CONTENT "+sDT.getContent());
//                //ToDo: The SDT may have traversable content too.
//            } else if (bodyElement instanceof XWPFTable) {
//                XWPFTable table = (XWPFTable)bodyElement;
//                System.out.println("TABLE"+table);
//                traverseTableRows(table.getRows());
//            }
//        }
//    }
//
//    public  void printDescriptionOfImagesInCell(XWPFDocument cell) {
//        List<XWPFParagraph> paragraphs = cell.getParagraphs();
//        for (XWPFParagraph paragraph : paragraphs) {
//            List<XWPFRun> runs = paragraph.getRuns();
//            for (XWPFRun run : runs) {
//                List<XWPFPicture> pictures = run.getEmbeddedPictures();
//                for (XWPFPicture picture : pictures) {
//                    //Do anything you want with the picture:
//                    System.out.println("Picture: " + picture.getDescription());
//                    System.out.println(picture.getPictureData());
//                }
//            }
//        }
//    }
//
//    public static String getMimeType(Context context, Uri uri) {
//        String extension;
//
//        if (uri.getScheme().equals(ContentResolver.SCHEME_CONTENT)) {
//            final MimeTypeMap mime = MimeTypeMap.getSingleton();
//            extension = mime.getExtensionFromMimeType(context.getContentResolver().getType(uri));
//        } else {
//            extension = MimeTypeMap.getFileExtensionFromUrl(Uri.fromFile(new File(uri.getPath())).toString());
//
//        }
//
//        return extension;
//    }
//
//    public void extractImages(XWPFDocument docx){
//        try {
//
//            //function to retrieve all the images from the doc file that have been selected
////
////            HWPFDocument wordDoc = new HWPFDocument(new FileInputStream(fileName));
////
//            XWPFWordExtractor extractor = new XWPFWordExtractor(docx);
//
////            extractText(extractor);
//
//            System.out.println(extractor.getText());
//
//            textView.setText(extractor.getText());
//            textView.setMovementMethod(new ScrollingMovementMethod());
//
//            int pages = docx.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();
//            Log.i("Pages",Integer.toString(pages));
//
////            List<XWPFPictureData> picList = docx.getAllPackagePictures();
////            int s = picList.size();
////            Log.i("Images",Integer.toString(s));
//
//            XWPFParagraph image = docx.createParagraph();
//            image.setAlignment(ParagraphAlignment.CENTER);
//
//            XWPFRun imageRun = image.createRun();
//            imageRun.setTextPosition(20);
//
//            List<XWPFPictureData> piclist = docx.getAllPictures();
//            int s = piclist.size();
//            Log.i("Images",Integer.toString(s));
//            // traverse through the list and write each image to a file
//            Iterator<XWPFPictureData> iterator = piclist.iterator();
//            int i = 0;
//            while (iterator.hasNext()) {
//                XWPFPictureData pic = iterator.next();
//                byte[] bytepic = pic.getData();
//                InputstreamGeneralPhotos1 = new ByteArrayInputStream(pic.getData());
//                imageRun.addPicture(openFileInput(pic.getFileName()),XWPFDocument.PICTURE_TYPE_PNG,pic.getFileName(),50,50);
//                if (InputstreamGeneralPhotos1 != null) {
//                    Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos1);
//                    img_generalPhotos1.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                }
//                i++;
//            }
//
////            for (XWPFPictureData pic : picList) {
////
////                System.out.print(pic.getPictureType());
////                System.out.print(pic.getData());
////                System.out.println("Image Number: " + i + " " + pic.getFileName());
////
////                switch (i){
////                    case 1:
////                        InputstreamGeneralPhotos1 = new ByteArrayInputStream(pic.getData());
////                        if (InputstreamGeneralPhotos1 != null){
////                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos1);
////                            img_generalPhotos1.setImageBitmap(selectedImage);
//////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
////                        }
////                        break;
////                    case 2:
////                        InputstreamGeneralPhotos2 = new ByteArrayInputStream(pic.getData());
////                        if (InputstreamGeneralPhotos2 != null){
////                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos2);
////                            img_generalPhotos2.setImageBitmap(selectedImage);
//////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
////                        }
////                        break;
////                    case 3:
////                        InputstreamGeneralPhotos3 = new ByteArrayInputStream(pic.getData());
////                        if (InputstreamGeneralPhotos3 != null){
////                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos3);
////                            img_generalPhotos3.setImageBitmap(selectedImage);
//////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
////                        }
////                        break;
////                    case 4:
////                        InputstreamGeneralPhotos4 = new ByteArrayInputStream(pic.getData());
////                        if (InputstreamGeneralPhotos4 != null){
////                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos4);
////                            img_generalPhotos4.setImageBitmap(selectedImage);
//////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
////                        }
////                        break;
////                    case 5:
////                        InputstreamGeneralPhotos5 = new ByteArrayInputStream(pic.getData());
////                        if (InputstreamGeneralPhotos5 != null){
////                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos5);
////                            img_generalPhotos5.setImageBitmap(selectedImage);
//////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
////                        }
////                        break;
////                    case 6:
////                        InputstreamGeneralPhotos6 = new ByteArrayInputStream(pic.getData());
////                        if (InputstreamGeneralPhotos6 != null){
////                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos6);
////                            img_generalPhotos6.setImageBitmap(selectedImage);
//////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
////                        }
////                        break;
////                    case 7:
////                        InputstreamBicyclePhotos1 = new ByteArrayInputStream(pic.getData());
////                        break;
////                    case 8:
////                        InputstreamBicyclePhotos2 = new ByteArrayInputStream(pic.getData());
////                        break;
////                    case 9:
////                        InputstreamFourhoursPhotos1 = new ByteArrayInputStream(pic.getData());
////                        break;
////                    case 10:
////                        InputstreamFourhoursPhotos2 = new ByteArrayInputStream(pic.getData());
////                        break;
////                    case 11:
////                        InputstreamFourhoursPhotos3 = new ByteArrayInputStream(pic.getData());
////                        break;
////                    case 12:
////                        InputstreamFourhoursPhotos4 = new ByteArrayInputStream(pic.getData());
////                        break;
////                }
////                i++;
////            }
//        } catch (Exception ex) {
//            ex.printStackTrace();
//        }
//
//    }
//
//    public  void addElements(XWPFPictureData pictureData,long w,long h){
//
//
//        ArrayList<ImageView> imageViews = new ArrayList<>();
////
//        ImageView image = new ImageView(this);
//        image.setLayoutParams(new RelativeLayout.LayoutParams((int)(w/2800),(int)(h/2800)));
//        image.setMaxHeight((int)h/2800);
//        image.setMaxWidth((int)w/2800);
////        image.setAdjustViewBounds(true);
////        image.setScaleType(ImageView.ScaleType.MATRIX);
//
//        InputStream inputStream = new ByteArrayInputStream(pictureData.getData());
//
//        Bitmap bitmap = BitmapFactory.decodeStream(inputStream);
//
////        Matrix matrix = new Matrix();
////        matrix.postScale(1/2800, 1/2800);
////
////        Bitmap scaledBitmap = Bitmap.createBitmap(bitmap,0,0,(int)w,(int)h,matrix,true);
//        image.setImageBitmap(bitmap);
//        mainUI.addView(image);
//
//
//        // Adds the view to the layout
//        LinearLayout imageLayout = new LinearLayout(getApplicationContext());
//        RelativeLayout.LayoutParams params = new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT, RelativeLayout.LayoutParams.WRAP_CONTENT);
//        params.addRule(RelativeLayout.BELOW, image.getId());
//        imageLayout.setLayoutParams(params);
//        mainUI.addView(imageLayout);
//
//        //Calling to add textview below image view
////        addTextViews();
//
////        for (int i = 0; i < 15; i++) {
//
////            final ImageView imageView = new ImageView(this);
////            InputStream inputStream = new ByteArrayInputStream(pictureData.getData());
////            Bitmap bitmap = BitmapFactory.decodeStream(inputStream);
////            imageView.setImageBitmap(bitmap);
////            Log.d("IMAGEDATA ",String.valueOf(bitmap));
//
////        }
//
//    }
//
////    addTextViews(run.toString(),run.getColor(),run.getFontFamily(),run.getFontSize());
//
//    public void addTextViews(String content, int s, Boolean b, UnderlinePatterns u){
//
//        TextView text = new TextView(this);
//        SpannableString c = new SpannableString(content);
//        c.setSpan(new UnderlineSpan(),0,content.length(),0);
//        text.setLayoutParams(new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT,RelativeLayout.LayoutParams.WRAP_CONTENT));
////        text.setBackgroundColor(Color.GRAY);
//
//       if(b) {
//           text.setTextColor(Color.BLACK);
//           text.setPadding(30, 10, 30, 10);
//           text.setTextSize((int) ((3 * s) / 2));
////           text.setTypeface(null, Typeface.BOLD);
//           text.setTextAppearance(R.style.FontFamily);
//
//           if(u == UnderlinePatterns.NONE){
//               text.setText(content);
//           }
//           else{
//               text.setText(c);
//           }
//           mainUI.addView(text);
//       }
//       else{
//           text.setTextColor(Color.BLACK);
//           text.setPadding(30, 10, 30, 10);
//           text.setTextSize((int) ((3 * s) / 2));
//
//           if(u == UnderlinePatterns.NONE){
//               text.setText(content);
//           }
//           else{
//               text.setText(c);
//           }
//           mainUI.addView(text);
//       }
//        // Adds the view to the layout
//        LinearLayout textLayout = new LinearLayout(getApplicationContext());
//        RelativeLayout.LayoutParams params = new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT, RelativeLayout.LayoutParams.WRAP_CONTENT);
//        params.addRule(RelativeLayout.BELOW, text.getId());
//        textLayout.setLayoutParams(params);
//        mainUI.addView(textLayout);
//
//
////        for (int i = 0; i < 15; i++) {
//
////            Log.d("ContentTEXT ",content);
////            final TextView textView = new TextView(this);
////            textView.setBackgroundColor(Color.GRAY);
////            textView.setTextColor(Color.BLUE);
////            textView.setTextSize(15);
////            textView.setText(content);
////            textLayout.addView(textView);
//
////        }
//    }
//
//    public void addElementsUI(String content, XWPFPictureData pictureData){
//        ImageView image = new ImageView(this);
//        image.setLayoutParams(new RelativeLayout.LayoutParams(200,200));
//        image.setMaxHeight(20);
//        image.setMaxWidth(20);
//        mainUI.addView(image);
//
//        TextView text = new TextView(this);
//        text.setLayoutParams(new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT,RelativeLayout.LayoutParams.WRAP_CONTENT));
//        text.setBackgroundColor(Color.GRAY);
//        text.setTextColor(Color.BLUE);
//        text.setTextSize(15);
//        mainUI.addView(text);
//
//        // Adds the view to the layout
//        LinearLayout Layouts = new LinearLayout(getApplicationContext());
//        RelativeLayout.LayoutParams params = new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT, RelativeLayout.LayoutParams.WRAP_CONTENT);
//
//        if(pictureData == null) {
//            params.addRule(RelativeLayout.BELOW, text.getId());
//            Layouts.setLayoutParams(params);
//            mainUI.addView(Layouts);
//
//            Log.d("ContentTEXT ",content);
//            final TextView textView = new TextView(this);
//            textView.setBackgroundColor(Color.GRAY);
//            textView.setTextColor(Color.BLUE);
//            textView.setTextSize(15);
//            textView.setText(content);
//            Layouts.addView(textView);
//
//        }else{
//            // Adds the view to the layout
//            params.addRule(RelativeLayout.BELOW, image.getId());
//            Layouts.setLayoutParams(params);
//            mainUI.addView(Layouts);
//
//            final ImageView imageView = new ImageView(this);
//            InputStream inputStream = new ByteArrayInputStream(pictureData.getData());
//            Bitmap bitmap = BitmapFactory.decodeStream(inputStream);
//            imageView.setImageBitmap(bitmap);
//            Log.d("IMAGEDATA ",String.valueOf(bitmap));
//        }
//
//    }
//
//
//
//    public void addTextViewsUsingLinearLayout(){
//        final String[] str = {"one","two","three","asdfgf"};
//        @SuppressLint("WrongViewCast")
//        RelativeLayout rl = (RelativeLayout) findViewById(R.id.MainUI);
//        final TextView[] tv = new TextView[paras.size()];
////        ArrayList<TextView> tv = new ArrayList<>();
////        for (int i=0; i<paras.size(); i++)
////        {
////            tv[i] = new TextView(this);
////            RelativeLayout.LayoutParams params=new RelativeLayout.LayoutParams
////                    ((int) RelativeLayout.LayoutParams.WRAP_CONTENT,(int) RelativeLayout.LayoutParams.WRAP_CONTENT);
////            params.leftMargin = 50;
////            params.topMargin  = i*50;
////            tv[i].setText(paras.get(i));
////            tv[i].setTextSize((float) 20);
////            tv[i].setPadding(20, 50, 20, 50);
////            tv[i].setLayoutParams(params);
////            rl.addView(tv[i]);
////
////
////        }
//
//
//        for (int i=0; i<str.length; i++)
//        {
//            tv[i] = new TextView(this);
//            RelativeLayout.LayoutParams params=new RelativeLayout.LayoutParams
//                    ((int) RelativeLayout.LayoutParams.WRAP_CONTENT,(int) RelativeLayout.LayoutParams.WRAP_CONTENT);
//            params.leftMargin = 50;
//            params.topMargin  = i*50;
//            tv[i].setText(str[i]);
//            tv[i].setTextSize((float) 20);
//            tv[i].setPadding(20, 50, 20, 50);
//            tv[i].setLayoutParams(params);
//            rl.addView(tv[i]);
//        }
//
//    }
//
//}


import android.view.Gravity;
public class MainActivity extends AppCompatActivity {
    LinearLayout mainUI;

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
    List<XWPFPictureData> picList;
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

        mainUI = findViewById(R.id.MainUI);

//        textView = (TextView)findViewById(R.id.textView);
//        img_generalPhotos1 = (ImageView)findViewById(R.id.imageView);
//        img_generalPhotos2 = (ImageView)findViewById(R.id.imageView2);

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
            if (resultCode == RESULT_OK) {
                switch (requestCode) {
                    case 111:
                        //this is action performed after openDocumentFromFileManager() when doc is selected
                        FileInputStream inputStream = (FileInputStream) getContentResolver().openInputStream(data.getData());

                        Uri uri = data.getData();
                        String extension = getMimeType(this, uri);
                        Log.i("Extension", extension);

                        XWPFDocument docx = new XWPFDocument(inputStream);
                        traverseBodyElements(docx.getBodyElements());
                        picList = docx.getAllPackagePictures();


                }
            }else {
                Toast.makeText(this, "Your file is not loaded", Toast.LENGTH_SHORT).show();
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void textViews(XWPFParagraph paragraph,List<IRunElement> runElements) {
        String paragraphText = paragraph.getParagraphText();

        int size;
        UnderlinePatterns u;
        Boolean b;
        String ff;

        for (IRunElement runElement : runElements) {
            if (runElement instanceof XWPFRun) {
                XWPFRun run = (XWPFRun) runElement;
                System.out.println("runClassName " + run.getClass().getName());
                System.out.println("run " + run);

                //Appending text to paragraph
//                paras.get(paraIndex).append(run);
                para.append(run);
                paras.add(para);
//                System.out.println("PARA1TEXT "+para.toString());
                Log.i("font family", run.getFontFamily());
                Log.i("font class", run.getFontName());

                size = run.getFontSize();
                u = run.getUnderline();
                b = run.isBold();
                ff = run.getFontFamily();

//                ArrayList<String> a = new ArrayList<String>();
//                a.add(paragraphText);
//
//
//                if(!a.contains(paragraphText)){
//                    a.add(paragraphText);
//                }

                if (paragraphText.length() > 1) {

                    addTextViews(paragraphText, size, b, u, ff);

                }


//                                addTextViews(paragrapthText);
//                                addElementsUI(paragrapthText,null);
//                    try {
////                        Thread.sleep(1000);
//
//                    } catch (InterruptedException e) {
//                        e.printStackTrace();
//                    }
//                addTextViews(run.toString(), run.getFontSize(), run.isBold(), run.getUnderline(), run.getFontFamily());



            }
        }

    }
    public void traversePictures(List<XWPFPicture> pictures)  {
        for (XWPFPicture picture : pictures) {

            //Create Image view to display images

//            Log.i("Desc",picture.getDescription());
            System.out.println("Picture "+picture);
            XWPFPictureData pictureData = picture.getPictureData();
            Log.i("PictureData ", pictureData.toString());

            long w = picture.getCTPicture().getSpPr().getXfrm().getExt().getCx();
            long h = picture.getCTPicture().getSpPr().getXfrm().getExt().getCy();

            DisplayMetrics displayMetrics = new DisplayMetrics();
            getWindowManager().getDefaultDisplay().getMetrics(displayMetrics);
            int height = displayMetrics.heightPixels;
            int width = displayMetrics.widthPixels;

            addElements(pictureData,w,h,height,width);
//            addElementsUI(null,pictureData);
        }

    }

    StringBuilder para = new StringBuilder();
    private static ArrayList<StringBuilder> paras = new ArrayList<>();
    private static int paraIndex = 0;

    public void traverseRunElements(List<IRunElement> runElements) throws Exception {

        System.out.println("PARAINDICES"+ paras.size());
        System.out.println("TRAVERSE RUN ELEMENTS");

        for (IRunElement runElement : runElements) {
            if (runElement instanceof XWPFFieldRun) {
                XWPFFieldRun fieldRun = (XWPFFieldRun)runElement;
                System.out.println("fieldRunClassName "+fieldRun.getClass().getName());
                System.out.println("fieldName "+fieldRun);
                traversePictures(fieldRun.getEmbeddedPictures());
            } else if (runElement instanceof XWPFHyperlinkRun) {
                XWPFHyperlinkRun hyperlinkRun = (XWPFHyperlinkRun)runElement;
                System.out.println("hyperLinkRunClassName "+ hyperlinkRun.getClass().getName());
                System.out.println("hyperlinkRun "+hyperlinkRun);
                traversePictures(hyperlinkRun.getEmbeddedPictures());
            } else if (runElement instanceof XWPFRun) {
                XWPFRun run = (XWPFRun)runElement;
                System.out.println("runClassName "+run.getClass().getName());
                System.out.println("run "+run);

                //Appending text to paragraph
//                paras.get(paraIndex).append(run);
                para.append(run);
                paras.add(para);
//                System.out.println("PARA1TEXT "+para.toString());
                Log.i("font family",run.getFontFamily());
                Log.i("font class",run.getFontName());

//                addTextViews(run.toString(),run.getFontSize(),run.isBold(),run.getUnderline(),run.getFontFamily());

                traversePictures(run.getEmbeddedPictures());

            } else if (runElement instanceof XWPFSDT) {
                XWPFSDT sDT = (XWPFSDT)runElement;
                System.out.println("sDT"+sDT);
                System.out.println("SDT_CONTENT "+sDT.getContent());
                //ToDo: The SDT may have traversable content too.
            }
        }
    }

    public void traverseTableCells(List<ICell> tableICells) throws Exception {
        for (ICell tableICell : tableICells) {
            if (tableICell instanceof XWPFSDTCell) {
                XWPFSDTCell sDTCell = (XWPFSDTCell)tableICell;
                System.out.println("sDTCELL "+sDTCell);
                //ToDo: The SDTCell may have traversable content too.
            } else if (tableICell instanceof XWPFTableCell) {
                XWPFTableCell tableCell = (XWPFTableCell)tableICell;
                System.out.println("TableCell "+tableCell);
                traverseBodyElements(tableCell.getBodyElements());
            }
        }
    }

    public void traverseTableRows(List<XWPFTableRow> tableRows) throws Exception {
        for (XWPFTableRow tableRow : tableRows) {
            System.out.println("TableRow "+tableRow);
            traverseTableCells(tableRow.getTableICells());
        }
    }

    public void traverseBodyElements(List<IBodyElement> bodyElements) throws Exception {
        System.out.println("TRAVERSE BODY ELEMENTS");

        for (IBodyElement bodyElement : bodyElements) {
            if (bodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph)bodyElement;
                System.out.println("PARA "+paragraph);

                //Creating textView & paragraph using String Builder
                paras.add(new StringBuilder());
                textViews(paragraph,paragraph.getIRuns());
                traverseRunElements(paragraph.getIRuns());
                paraIndex = paraIndex + 1;

            } else if (bodyElement instanceof XWPFSDT) {
                XWPFSDT sDT = (XWPFSDT)bodyElement;
                System.out.println("SDT"+sDT);
                System.out.println("SDT_CONTENT "+sDT.getContent());
                //ToDo: The SDT may have traversable content too.
            } else if (bodyElement instanceof XWPFTable) {
                XWPFTable table = (XWPFTable)bodyElement;
                System.out.println("TABLE"+table);
                traverseTableRows(table.getRows());
            }
        }
    }

    public  void printDescriptionOfImagesInCell(XWPFDocument cell) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                List<XWPFPicture> pictures = run.getEmbeddedPictures();
                for (XWPFPicture picture : pictures) {
                    //Do anything you want with the picture:
                    System.out.println("Picture: " + picture.getDescription());
                    System.out.println(picture.getPictureData());
                }
            }
        }
    }

    public static String getMimeType(Context context, Uri uri) {
        String extension;

        if (uri.getScheme().equals(ContentResolver.SCHEME_CONTENT)) {
            final MimeTypeMap mime = MimeTypeMap.getSingleton();
            extension = mime.getExtensionFromMimeType(context.getContentResolver().getType(uri));
        } else {
            extension = MimeTypeMap.getFileExtensionFromUrl(Uri.fromFile(new File(uri.getPath())).toString());

        }

        return extension;
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

//            List<XWPFPictureData> picList = docx.getAllPackagePictures();
//            int s = picList.size();
//            Log.i("Images",Integer.toString(s));

            XWPFParagraph image = docx.createParagraph();
            image.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun imageRun = image.createRun();
            imageRun.setTextPosition(20);

            List<XWPFPictureData> piclist = docx.getAllPictures();
            int s = piclist.size();
            Log.i("Images",Integer.toString(s));
            // traverse through the list and write each image to a file
            Iterator<XWPFPictureData> iterator = piclist.iterator();
            int i = 0;
            while (iterator.hasNext()) {
                XWPFPictureData pic = iterator.next();
                byte[] bytepic = pic.getData();
                InputstreamGeneralPhotos1 = new ByteArrayInputStream(pic.getData());
                imageRun.addPicture(openFileInput(pic.getFileName()),XWPFDocument.PICTURE_TYPE_PNG,pic.getFileName(),50,50);
                if (InputstreamGeneralPhotos1 != null) {
                    Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos1);
                    img_generalPhotos1.setImageBitmap(selectedImage);
//                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
                }
                i++;
            }
//            for (XWPFPictureData pic : picList) {
//
//                System.out.print(pic.getPictureType());
//                System.out.print(pic.getData());
//                System.out.println("Image Number: " + i + " " + pic.getFileName());
//
//                switch (i){
//                    case 1:
//                        InputstreamGeneralPhotos1 = new ByteArrayInputStream(pic.getData());
//                        if (InputstreamGeneralPhotos1 != null){
//                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos1);
//                            img_generalPhotos1.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                        }
//                        break;
//                    case 2:
//                        InputstreamGeneralPhotos2 = new ByteArrayInputStream(pic.getData());
//                        if (InputstreamGeneralPhotos2 != null){
//                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos2);
//                            img_generalPhotos2.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                        }
//                        break;
//                    case 3:
//                        InputstreamGeneralPhotos3 = new ByteArrayInputStream(pic.getData());
//                        if (InputstreamGeneralPhotos3 != null){
//                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos3);
//                            img_generalPhotos3.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                        }
//                        break;
//                    case 4:
//                        InputstreamGeneralPhotos4 = new ByteArrayInputStream(pic.getData());
//                        if (InputstreamGeneralPhotos4 != null){
//                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos4);
//                            img_generalPhotos4.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                        }
//                        break;
//                    case 5:
//                        InputstreamGeneralPhotos5 = new ByteArrayInputStream(pic.getData());
//                        if (InputstreamGeneralPhotos5 != null){
//                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos5);
//                            img_generalPhotos5.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                        }
//                        break;
//                    case 6:
//                        InputstreamGeneralPhotos6 = new ByteArrayInputStream(pic.getData());
//                        if (InputstreamGeneralPhotos6 != null){
//                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos6);
//                            img_generalPhotos6.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                        }
//                        break;
//                    case 7:
//                        InputstreamBicyclePhotos1 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 8:
//                        InputstreamBicyclePhotos2 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 9:
//                        InputstreamFourhoursPhotos1 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 10:
//                        InputstreamFourhoursPhotos2 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 11:
//                        InputstreamFourhoursPhotos3 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 12:
//                        InputstreamFourhoursPhotos4 = new ByteArrayInputStream(pic.getData());
//                        break;
//                }
//                i++;
//            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public  void addElements(XWPFPictureData pictureData,long w,long h,int height,int width){

        ArrayList<ImageView> imageViews = new ArrayList<>();
//
        ImageView image = new ImageView(this);

        if((int)w > width){
            w = width - 100;
            h = h / (3 *(width-100));
        }

        image.setForegroundGravity(Gravity.CENTER);
        image.setPadding(100,30,50,30);
        image.setLayoutParams(new RelativeLayout.LayoutParams((int)(w),(int)(h)));
        image.setMaxHeight((int)h/2800);
        image.setMaxWidth((int)w/2800);
//        image.setAdjustViewBounds(true);
//        image.setScaleType(ImageView.ScaleType.MATRIX);
        InputStream inputStream = new ByteArrayInputStream(pictureData.getData());
        Bitmap bitmap = BitmapFactory.decodeStream(inputStream);

//        Matrix matrix = new Matrix();
//        matrix.postScale(1/2800, 1/2800);
//        Bitmap scaledBitmap = Bitmap.createBitmap(bitmap,0,0,(int)w,(int)h,matrix,true);

        image.setImageBitmap(bitmap);
        mainUI.addView(image);


        // Adds the view to the layout
        LinearLayout imageLayout = new LinearLayout(getApplicationContext());
        RelativeLayout.LayoutParams params = new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT, RelativeLayout.LayoutParams.WRAP_CONTENT);
        params.addRule(RelativeLayout.BELOW, image.getId());
        imageLayout.setLayoutParams(params);
        mainUI.addView(imageLayout);

        //Calling to add textview below image view
//        addTextViews();

//        for (int i = 0; i < 15; i++) {
//            final ImageView imageView = new ImageView(this);
//            InputStream inputStream = new ByteArrayInputStream(pictureData.getData());
//            Bitmap bitmap = BitmapFactory.decodeStream(inputStream);
//            imageView.setImageBitmap(bitmap);
//            Log.d("IMAGEDATA ",String.valueOf(bitmap));
//        }
    }


    public static int stringCompare(String str1, String str2)
    {

        int l1 = str1.length();
        int l2 = str2.length();
        int lmin = Math.min(l1, l2);

        for (int i = 0; i < lmin; i++) {
            int str1_ch = (int)str1.charAt(i);
            int str2_ch = (int)str2.charAt(i);

            if (str1_ch != str2_ch) {
                return str1_ch - str2_ch;
            }
        }

        // Edge case for strings like
        // String 1="Geeks" and String 2="Geeksforgeeks"
        if (l1 != l2) {
            return l1 - l2;
        }

        // If none of the above conditions is true,
        // it implies both the strings are equal
        else {
            return 0;
        }
    }
    int TagCnt = 0;
    public void addTextViews(String content, int s, Boolean b, UnderlinePatterns u,String f){


        TextView text = new TextView(this);
        text.setLayoutParams(new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT,RelativeLayout.LayoutParams.WRAP_CONTENT));
        SpannableString c = new SpannableString(content);
        c.setSpan(new UnderlineSpan(),0,content.length(),0);
        text.setTag(TagCnt);

        ArrayList<TextView> textViews = new ArrayList<>();
//        TextView myText = (TextView) mainUI.findViewWithTag(TagCnt - i); // get the element

        for (int i=TagCnt - 1;i >= 0;i--){

            TextView myText = (TextView) mainUI.findViewWithTag(i); // get the element
            if(myText != null) {
                Log.d("CONT ",String.valueOf(content));
                if((content.equals(myText.getText().toString()))){
                    Log.d("MYTXT ",myText.getText().toString());
                    Log.d("ICNT ",String.valueOf(i));
                    break;

                }else{
                    setProperText(text, content, s, b, u, f);
                    break;
                }
            }else{
                Log.d("NTAGCNT ",String.valueOf(TagCnt));
                Log.d("NICNT ",String.valueOf(i));
                Log.d("NULLTXT ",content);
            }
        }
        if(TagCnt == 0 ){
            setProperText(text, content, s, b, u, f);
        }
        TagCnt = TagCnt + 1;

    }

    public void setProperText(TextView text,String content, int s, Boolean b, UnderlinePatterns u,String fontFamily){
        SpannableString c = new SpannableString(content);
        c.setSpan(new UnderlineSpan(),0,content.length(),0);
        text.setPadding(60, 10, 50, 10);

        if (b) {
            text.setTextColor(Color.BLACK);
            text.setTextSize((int) ((3 * s) / 2));
            //text.setTypeface(null, FontStyle.fontFamily);
            text.setTypeface(null, Typeface.BOLD);

            if (u == UnderlinePatterns.NONE) {
                text.setText(content);
            } else {
                //text.setTextAlignment(View.TEXT_ALIGNMENT_CENTER);
                text.setText(c);
                text.setGravity(Gravity.CENTER);
            }
//                if(!c.equals(myText.getText().toString()) || !content.equals(myText.getText().toString())) {
            mainUI.addView(text);
//                }
        } else {
            text.setTextColor(Color.BLACK);
            text.setTextSize((int) ((3 * s) / 2));

            if (u == UnderlinePatterns.NONE) {
                text.setText(content);
            } else {
                //               text.setTextAlignment(View.TEXT_ALIGNMENT_CENTER);
                text.setText(c);
                text.setGravity(Gravity.CENTER);
            }
//                if(!c.equals(myText.getText().toString()) || !content.equals(myText.getText().toString())) {
            mainUI.addView(text);
//                }
        }


        // Adds the view to the layout
        LinearLayout textLayout = new LinearLayout(getApplicationContext());
        RelativeLayout.LayoutParams params = new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT, RelativeLayout.LayoutParams.WRAP_CONTENT);
        params.addRule(RelativeLayout.BELOW, text.getId());
        textLayout.setLayoutParams(params);
        mainUI.addView(textLayout);


    }

    public void addElementsUI(String content, XWPFPictureData pictureData){
        ImageView image = new ImageView(this);
        image.setLayoutParams(new RelativeLayout.LayoutParams(200,200));
        image.setMaxHeight(20);
        image.setMaxWidth(20);
        mainUI.addView(image);

        TextView text = new TextView(this);
        text.setLayoutParams(new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT,RelativeLayout.LayoutParams.WRAP_CONTENT));
        text.setBackgroundColor(Color.GRAY);
        text.setTextColor(Color.BLUE);
        text.setTextSize(15);
        mainUI.addView(text);

        // Adds the view to the layout
        LinearLayout Layouts = new LinearLayout(getApplicationContext());
        RelativeLayout.LayoutParams params = new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT, RelativeLayout.LayoutParams.WRAP_CONTENT);

        if(pictureData == null) {
            params.addRule(RelativeLayout.BELOW, text.getId());
            Layouts.setLayoutParams(params);
            mainUI.addView(Layouts);

            Log.d("ContentTEXT ",content);
            final TextView textView = new TextView(this);
            textView.setBackgroundColor(Color.GRAY);
            textView.setTextColor(Color.BLUE);
            textView.setTextSize(15);
            textView.setText(content);
            Layouts.addView(textView);

        }else{
            // Adds the view to the layout
            params.addRule(RelativeLayout.BELOW, image.getId());
            Layouts.setLayoutParams(params);
            mainUI.addView(Layouts);

            final ImageView imageView = new ImageView(this);
            InputStream inputStream = new ByteArrayInputStream(pictureData.getData());
            Bitmap bitmap = BitmapFactory.decodeStream(inputStream);
            imageView.setImageBitmap(bitmap);
            Log.d("IMAGEDATA ",String.valueOf(bitmap));
        }

    }



    public void addTextViewsUsingLinearLayout(){
        final String[] str = {"one","two","three","asdfgf"};
        @SuppressLint("WrongViewCast")
        RelativeLayout rl = (RelativeLayout) findViewById(R.id.MainUI);
        final TextView[] tv = new TextView[paras.size()];
//        ArrayList<TextView> tv = new ArrayList<>();
//        for (int i=0; i<paras.size(); i++)
//        {
//            tv[i] = new TextView(this);
//            RelativeLayout.LayoutParams params=new RelativeLayout.LayoutParams
//                    ((int) RelativeLayout.LayoutParams.WRAP_CONTENT,(int) RelativeLayout.LayoutParams.WRAP_CONTENT);
//            params.leftMargin = 50;
//            params.topMargin  = i*50;
//            tv[i].setText(paras.get(i));
//            tv[i].setTextSize((float) 20);
//            tv[i].setPadding(20, 50, 20, 50);
//            tv[i].setLayoutParams(params);
//            rl.addView(tv[i]);
//
//
//        }


        for (int i=0; i<str.length; i++)
        {
            tv[i] = new TextView(this);
            RelativeLayout.LayoutParams params=new RelativeLayout.LayoutParams
                    ((int) RelativeLayout.LayoutParams.WRAP_CONTENT,(int) RelativeLayout.LayoutParams.WRAP_CONTENT);
            params.leftMargin = 50;
            params.topMargin  = i*50;
            tv[i].setText(str[i]);
            tv[i].setTextSize((float) 20);
            tv[i].setPadding(20, 50, 20, 50);
            tv[i].setLayoutParams(params);
            rl.addView(tv[i]);
        }

    }

}