package com.jdreamer;

import net.xdevelop.jpclient.PyResult;
import net.xdevelop.jpclient.PyServeContext;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Method;
import java.time.LocalDateTime;


class Style {
    private Double fontSize;
    private String fontFamily;

    Style(Double fontSize, String fontFamily) {
        this.fontSize = fontSize;

        this.fontFamily = fontFamily;
    }

    public Double getFontSize() {
        return fontSize;
    }

    public String getFontFamily() {
        return fontFamily;
    }
}

public class PowerPointTranslate {
    public static void main(String[] args) throws Exception {
        String fileName = "TODO";

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(fileName));

        // init the PyServeContext, it will make a connection to JPServe
        /*
        >>> from jpserve.jpserve import JPServe
        >>> serve = JPServe(("localhost", 8888))
        >>> serve.start()
         */
        PyServeContext.init("localhost", 8888);

        System.out.println("Started at: " + LocalDateTime.now());
        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape sh : slide.getShapes()) {
                // shapeName of the shape
                String shapeName = sh.getShapeName();

                if (sh instanceof XSLFTextShape) {
                    XSLFTextShape shape = (XSLFTextShape) sh;

                    for (XSLFTextParagraph para : shape.getTextParagraphs()) {
                        String text = para.getText().trim();

                        Style style = null;

                        if (!text.isEmpty()) {
                            try {
                                if (!para.getTextRuns().isEmpty()) {
                                    XSLFTextRun firstTextRun = para.getTextRuns().get(0);

                                    style = new Style(firstTextRun.getFontSize(), firstTextRun.getFontFamily());

                                    Method m = para.getClass().getDeclaredMethod("clearButKeepProperties", null);
                                    m.setAccessible(true);
                                    m.invoke(para, null);
                                }


                                XSLFTextRun textRun = para.addNewTextRun();
                                if (style != null) {
                                    textRun.setFontSize(style.getFontSize() - 3);
                                    textRun.setFontFamily(style.getFontFamily());

                                    if (!shapeName.contains("タイトル")) {
                                        textRun.setFontColor(Color.BLACK);
                                    }
                                }

                                System.out.println(text);

                                String translated = translate(text);
                                if (!translated.isEmpty()) {
                                    textRun.setText(translated.substring(1, translated.length() - 1));
                                }
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                    }
                }
            }
        }

        ppt.write(new FileOutputStream("translated.pptx"));

        System.out.println("End at: " + LocalDateTime.now());
    }

    private static String translate(String text) {
        // prepare the script, and assign the return value to _result_
        String script = "#-*- coding: utf-8 -*-\n"
                + "from googletrans import Translator\n"
                + "translator = Translator()\n"
                + "_result_ = translator.translate(u\"\"\"" + text + "\"\"\", dest='en').text";

        //System.out.println(script);

        PyResult rs = PyServeContext.getExecutor().exec(script);

        if (rs.isSuccess()) {
            return rs.getResult();
        } else {
            System.out.println("Execute python script failed: " + rs.getMsg());

            return "";
        }
    }
}
