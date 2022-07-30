package com.zwnop.wps.dispimg.handle;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import org.apache.commons.lang3.StringUtils;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class Test {
    private static final String FILE_PATH = "src/main/resources/test.xlsx";
    private static final String OUTPUT_DIR = "D:\\student\\";

    public static void main(String[] args) throws Exception {
        File file = new File(FILE_PATH);
        List<Student> students = ExcelImportUtil.importExcel(new FileInputStream(file), Student.class, new ImportParams());

        ZipFile zipFile = new ZipFile(file);
        Document imgRelDoc = getDoc(zipFile, zipFile.getEntry("xl/_rels/cellimages.xml.rels"));
        Map<String, String> idPathMap = handleImgRelDoc(imgRelDoc);

        Document imgDoc = getDoc(zipFile, zipFile.getEntry("xl/cellimages.xml"));
        Map<String, String> randomIdMap = handleImgDoc(imgDoc);

        Map<String, String> randomPathMap = new HashMap<>();
        for (Map.Entry<String, String> entry : randomIdMap.entrySet()) {
            randomPathMap.put(entry.getKey(), idPathMap.get(entry.getValue()));
        }
        outputImg(zipFile, students, randomPathMap);
        for (Student student : students) {
            System.out.println(student);
        }
    }

    private static void outputImg(ZipFile zf, List<Student> students, Map<String, String> randomPathMap) throws IOException {
        String regex = ".DISPIMG\\(\\\"([A-Za-z0-9_]*)\\\".*";
        Map<String, String> randomTypeMap = new HashMap<>();
        Map<String, String> randomNameTypeMap = new HashMap<>();
        for (Map.Entry<String, String> entry : randomPathMap.entrySet()) {
            randomTypeMap.put(entry.getKey(), getFileType(entry.getValue()));
            randomNameTypeMap.put(entry.getKey(), getFileNameType(entry.getValue()));
        }

        for (Student student : students) {
            String icon = student.getIcon();
            if (StringUtils.isBlank(icon)) {
                continue;
            }
            String random = icon.replaceAll(regex, "$1");
            String randomFileName = random + "." + randomTypeMap.get(random);
            String outputPath = OUTPUT_DIR + randomFileName;
            student.setIcon(outputPath);
            String zipPath = "xl/media/" + randomNameTypeMap.get(random);
            ZipEntry img = zf.getEntry(zipPath);
            write(zf, img, outputPath);
        }
    }

    private static void write(ZipFile zf, ZipEntry img, String outputPath) throws IOException {
        new File(outputPath).createNewFile();
        try(FileOutputStream os = new FileOutputStream(outputPath);
            InputStream is = zf.getInputStream(img)){
            byte[] buff = new byte[1024];
            int len = 0;
            while ((len = is.read(buff)) != -1) {
                os.write(buff, 0, len);
            }
        }
    }

    private static Map<String, String> handleImgDoc(Document doc) {
        Map<String, String> map = new HashMap<>();
        List<Element> imgs = doc.getRootElement().elements("cellImage");
        for (Element img : imgs) {
            Element pic = img.element("pic");
            Element nvPicPr = pic.element("nvPicPr");
            Element cNvPr = nvPicPr.element("cNvPr");
            String random = cNvPr.attributeValue("name");

            Element blipFill = pic.element("blipFill");
            Element blip = blipFill.element("blip");
            String rId = blip.attributeValue("embed");
            map.put(random, rId);
        }
        return map;
    }

    private static Map<String, String> handleImgRelDoc(Document doc){
        Map<String, String> map = new HashMap<>();
        Element root = doc.getRootElement();
        List<Element> elements = root.elements("Relationship");
        for (Element element : elements) {
            map.put(element.attributeValue("Id"), element.attributeValue("Target"));
        }
        return map;
    }

    private static Document getDoc(ZipFile zf, ZipEntry entry) throws Exception {
        SAXReader saxReader = new SAXReader();
        InputStream is = zf.getInputStream(entry);
        return saxReader.read(is);
    }

    private static String getFileName(String dirName) {
        String regex = "^media/([A-Za-z0-9_]*)\\..*";
        return dirName.replaceAll(regex, "$1");
    }

    private static String getFileNameType(String dirName) {
        String regex = "^media/([A-Za-z0-9_\\.]*)$";
        return dirName.replaceAll(regex, "$1");
    }

    private static String getFileType(String dirName) {
        String regex = ".*\\.([A-Za-z0-9_]*)$";
        return dirName.replaceAll(regex, "$1");
    }
}
