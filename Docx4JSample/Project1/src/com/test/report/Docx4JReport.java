package com.test.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.StringTokenizer;

import javax.xml.bind.JAXBElement;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.Text;



public class Docx4JReport {
    
    private static String stag="«";
    private static String etag="»";
    
    public Docx4JReport() {
        super();
    }

    public static String mailMergeTemplate(String templateDoc, String outputFileName,
                                    String picLocation) {

        try {
            WordprocessingMLPackage template =
                WordprocessingMLPackage.load(new File(templateDoc));

            MainDocumentPart mdp = template.getMainDocumentPart();
            List<Object> runs = getAllElementFromObject(mdp, Text.class);
            for (Object text : runs) {
                Text textElement = (Text)text;
                System.out.println(textElement.getValue());
                if (textElement.getValue().equalsIgnoreCase("«name»")) {
                    textElement.setValue("Gopinath Jayavel");
                }
                if (textElement.getValue().equalsIgnoreCase("«addressline1»")) {
                    textElement.setValue("03 407 Dubai Cross Street");
                }
                if (textElement.getValue().equalsIgnoreCase("«addressline2»")) {
                    textElement.setValue("Dubai Main Road");
                }
                if (textElement.getValue().equalsIgnoreCase("«addressline3»")) {
                    textElement.setValue("Dubai");
                }
                if (textElement.getValue().equalsIgnoreCase("«addressline4»")) {
                    textElement.setValue("UAE");
                }
            }

            Parts parts = template.getParts();

            HashMap partsMap = parts.getParts();
            PartName partName = null;
            Part part = null;

            Set set = partsMap.keySet();
            for (Iterator iterator = set.iterator(); iterator.hasNext(); ) {
                PartName name1 = (PartName)iterator.next();
                System.out.println(" NAME :" + name1.getName());
                if (name1.getName().equalsIgnoreCase("/word/media/image1.jpeg")) {
                    part = (Part)partsMap.get(name1);
                    partName = name1;
                }

            }
            if (part != null && partName != null) {
                part = (Part)partsMap.get(partName);
                BinaryPart binaryPart = (BinaryPart)part;
                try {
                    binaryPart.setBinaryData(fileToBytes(new File(picLocation)));
                } catch (FileNotFoundException e) {
                } catch (IOException e) {
                }
            }


            template.save(new File(outputFileName));
            System.out.println("Done");
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("Errors");
        }

        return "SUCCESS";
    }


    /**
     * This method returns all the placeholder with value get(0) is the place holder attributed seperated with (,) and values in get(1).
     * @param contentMap
     * @param templateFile
     * @param outPutFile
     * @return
     */
    public static String mailMergeCsv(Map contentMap,String templateFile, String outPutFile) {
        String status = "FAILURE";
        try {
            
            WordprocessingMLPackage template =
                            WordprocessingMLPackage.load(new File(templateFile));
      
            MainDocumentPart mdp = template.getMainDocumentPart();
            List<Object> runs = getAllElementFromObject(mdp, Text.class);
        
                Map map1 = mailMergeCreateData((String)contentMap.get(0), (String)contentMap.get(1));
                contentMap.putAll(map1);
                mailMergeParagraph(runs,mdp,contentMap);
                    
            template.save(new File(outPutFile));
            status = "SUCCESS";
            System.out.println("Done");
        } catch (Exception e) {
            e.printStackTrace();
            status = "FILE_NOT_FOUND : Template file " + templateFile +" is missing or Access Denied.";
        }

        return status;
    }
    
    
    public static String mailMergeParagraph(List<Object> runs, MainDocumentPart mdp, Map contentMap) {

        try {

            String content1="";
            String content = (String)contentMap.get("«content_p»");
            
            if(content != null){
                content1 = content.replace("...", "newline");// the content which is written line by line which is separated with "..."
                content1 = content1.replace("\n", "");
                content1 = content1.replace("\t", "    ");//tab or carriage return should be handled here. else the spacing wil not replace in template.
            }
           
            
            for (Object text : runs) {
                Text textElement = (Text)text;
                 System.out.println(textElement.getValue());
                if(textElement.getValue().contains("«")){
                     if (textElement.getValue().equalsIgnoreCase("«content_p»")) {
                           writeAsParagraph(mdp, content1, "«content_p»", "newline");
                    }else if(contentMap.get(textElement.getValue()) != null){
                            String k = (String)contentMap.get(textElement.getValue());//Map key and tempalte place holder will be same.
                            textElement.setSpace("preserve");
                            textElement.setValue(k);
                     }
                } 
            }
            

        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("Errors");
        }

        return "SUCCESS";
    }
    
    
    /**
     * This method parses the comma(,) separated values with matching value. The placeholder attribute then created with tags '«' '»' and value as value in the same position.
     * @param header
     * @param data
     * @return map with («placeholder», value) - which is easy to iterate with template document to repalce.
     */
    
    public static Map mailMergeCreateData(String header, String data) {
        HashMap map = null;
        try {
            
            if(header != null && data != null){
                
                String[] headerArr = header.replaceAll("\"", "").split(",");
                String[] dataArr = data.replaceAll("\"", "").split(",");
                int i=0;
                int length = dataArr.length;
                //String tempData="";
                String tempkey="";
                map = new HashMap();
                    for(String s : headerArr){
                        if(s!= null && i<length){
                            tempkey  = stag+s.trim()+etag;
                            //System.out.println(tempkey +" , "+dataArr[i]);
                             map.put(tempkey, dataArr[i]);
                        }
                
                        i++;
                
                    }  
                    
                System.out.println("===========================");
            }
            
        }catch(Exception e){
            e.printStackTrace();
        }
        
        return map;
    }
    
    /**
     *
     * @param mdp
     * @param content - content seperated with a splitter like , or ... or ## or some spcial characters
     * @param palceholder
     * @param splitStr - content to split and write line by line
     */
    public static void writeAsParagraph(MainDocumentPart mdp,String content, String palceholder,String splitStr){
        List<Object> paragraphs =  getAllElementFromObject(mdp, P.class);
        P toReplace = null;
        for(Object p:paragraphs){
            List<Object> texts = getAllElementFromObject(p, Text.class);
            for(Object t:texts){
                Text contents = (Text)t;
                if(contents.getValue() != null && contents.getValue().equalsIgnoreCase(palceholder)){
                    toReplace = (P)p;
                    break;
                }
            }
        }
         StringTokenizer st = new StringTokenizer(content,splitStr);
         String ptext = null;
         P copy = null;
         List<?> texts = null;
        while (st.hasMoreTokens()){
             ptext = (String)st.nextToken();
            copy = (P)XmlUtils.deepCopy(toReplace);
            texts = getAllElementFromObject(copy, Text.class);
            if(texts.size() >0){
                Text textToReplace = (Text) texts.get(1);
                textToReplace.setSpace("preserve");
                textToReplace.setValue(ptext.trim());
            }
            
            ((ContentAccessor)toReplace.getParent()).getContent().add(copy);
        }
        
        ((ContentAccessor)toReplace.getParent()).getContent().remove(toReplace);
    }
    
    private static List<Object> getAllElementFromObject(Object obj,
                                                        Class<?> toSearch) {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement)
            obj = ((JAXBElement<?>)obj).getValue();

        if (obj != null && obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor)obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }

        }
        return result;
    }
    
    private static byte[] fileToBytes(File file) throws FileNotFoundException,
                                                        IOException {
        byte[] bytes = null;
        // Our utility method wants that as a byte array
        if (file.exists() && file.isFile()) {
            java.io.InputStream is = new java.io.FileInputStream(file);
            long length = file.length();
            // You cannot create an array using a long type.
            // It needs to be an int type.
            bytes = new byte[(int)length];
            int offset = 0;
            int numRead = 0;
            while (offset < bytes.length &&
                   (numRead = is.read(bytes, offset, bytes.length - offset)) >=
                   0) {
                offset += numRead;
            }
            // Ensure all the bytes have been read in
            if (offset < bytes.length) {
                // System.out.println("Could not completely read file
                // "+file.getName());
            }
            is.close();
        } else {
            bytes = new byte[0];
        }
        return bytes;
    }
    
}
