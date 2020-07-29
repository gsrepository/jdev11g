package com.test.report;

import java.util.HashMap;
import java.util.Map;

public class TestMe {
    public TestMe() {
        super();
    }
    
    public static void main(String[] args) {
        String templateDoc1 = "C:\\Oracle\\JdevHome\\mywork\\Docx4JSample\\Project1\\src\\template\\SimpleReport.docx";
        String outputFileName1 = "C:\\Oracle\\JdevHome\\mywork\\Docx4JSample\\outptsaved\\SimpleReport_output.docx";
        String picLocation ="C:\\Oracle\\JdevHome\\mywork\\panda.jpg";
        
        //Docx4JReport.mailMergeTemplate(templateDoc1, outputFileName1, picLocation);
        
               
        String templateDoc2 = "C:\\Oracle\\JdevHome\\mywork\\Docx4JSample\\Project1\\src\\template\\CaseReport.docx";
        String outputFileName2 = "C:\\Oracle\\JdevHome\\mywork\\Docx4JSample\\outptsaved\\CaseReport_output.docx";
        
        
        Map contentMap = new HashMap();
        contentMap.put(0,"\"created_by\",\"rpt_date\",\"date\",\"submitBy\",\"date\",\"content_p\",\"comments\"");
        contentMap.put(1,"\"Gopinath\",\"24 JUL 2020\",\"24 JUL 2020\",\"Sharad\",\"25 JUL 2020\",\"140037F	NAME OF 140037F                                   	ABDC01	123456-A1	90	(1.0%)...	\n" + 
        "100051H	NAME OF 100051H                                   	ABDC01	123456-AC	90	(1.6%)...	\n" + 
        "100000G	NAME OF 100000G                                   	ABDC01	123456-AB	90	(2.1%)...	\n" + 
        "101234Y	NAME OF 101234Y                                   	ABDC01	123456-A4	90	(2.6%)...	\n" + 
        "101234G	NAME OF 101234G                                   	ABDC01	123456-A4	90	(3.1%)...	\n" + 
        "101234B	NAME OF 101234B                                   	ABDC01	123456-A1	80	(3.6%)...	\n" + 
        "101234G	NAME OF 101234G                                   	ABDC01	123456-AA	80	(4.1%)...	\",\"recommend for scholarship\"");
        
        Docx4JReport.mailMergeCsv(contentMap, templateDoc2, outputFileName2);
    }
}
