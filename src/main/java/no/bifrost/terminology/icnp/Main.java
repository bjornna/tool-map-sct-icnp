package no.bifrost.terminology.icnp;

import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.SneakyThrows;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

public class Main {
    static  String ROOT = "LOADED_FROM config.properties";
    static String ICNP = "norsk_icnp_2019.xlsx";
    static String SCT_DIAGNOSIS = "ICNP-SCT Mapping-Diagnoses 2019.xlsx";
    static  String SCT_INTERVENTIONS = "ICNP-SCT Mapping-Interventions 2019.xlsx";
    static final String EXPORT_JSON_FILE = "icnp.json";
    static final String EXPORT_EXCEL_FILE = "icnp-snomedct.xlsx";

    @SneakyThrows
    static void LoadProperties(){
        FileInputStream fis = new FileInputStream(new File("config.properties"));
        Properties properties = new Properties();
        properties.load(fis);
        ROOT = properties.getProperty("root");
        System.out.println("ROOT is " + ROOT);


    }
    public static void main(String[] args) {
        LoadProperties();
        DoWork();

    }
    static void DoWork(){

        var map = loadICNP();
        map = addSnomedTerm(new File(ROOT, SCT_DIAGNOSIS), map);
        map = addSnomedTerm(new File(ROOT, SCT_INTERVENTIONS), map);

        mapToJson(EXPORT_JSON_FILE, map);
        try {
            writeToExcel(EXPORT_EXCEL_FILE, map);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    static String createSnomedUrl(String code){
        return "https://browser.ihtsdotools.org/?perspective=full&conceptId1=" + code +  "&edition=MAIN/SNOMEDCT-NO/2020-10-15&release=&languages=no,en";
    }
    static String createICNPUrl(String code){
        return "https://neuronsong.com//_/_sites/icnp-browser/#/2019/concepts/no/" + code;
    }
    static void writeToExcel(String filename, Map<String,ICNPTerm> terms) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("ICNP-SCT");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("ICNP");
        header.createCell(1).setCellValue("AXIS");
        header.createCell(2).setCellValue("SCT");
        header.createCell(3).setCellValue("TERM");
        header.createCell(4).setCellValue("DEFINITION");
        int row = 1;
        for(ICNPTerm t: terms.values()){
            Row r = sheet.createRow(row);
            var icnpCell = r.createCell(0);
            icnpCell.setCellValue(t.getCode());
            Hyperlink icnpHyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
            icnpHyperlink.setAddress(createICNPUrl(t.getCode()));
            icnpHyperlink.setLabel(t.getCode());
            icnpCell.setCellValue(t.getCode());
            icnpCell.setHyperlink(icnpHyperlink);

            r.createCell(1).setCellValue(t.getAxis());
            if(t.getSnomedTerm() != null){
                var url = createSnomedUrl(t.getSnomedTerm());
                var cell = r.createCell(2);
                Hyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
                hyperlink.setAddress(url);
                hyperlink.setLabel(t.getSnomedTerm());
                cell.setCellValue(t.getSnomedTerm());
                cell.setHyperlink(hyperlink);
            }

            r.createCell(3).setCellValue(t.getTerm());
            r.createCell(4).setCellValue(t.getDefinition());
            row++;
        }
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) +filename;
        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();
    }

    @SneakyThrows
    static void mapToJson(String filename, Map<String, ICNPTerm> terms) {
        ObjectMapper mapper = new ObjectMapper();
        var values = terms.values();
        mapper.writerWithDefaultPrettyPrinter().writeValue(new File(filename), values);
    }

    static Map<String, ICNPTerm> addSnomedTerm(File mapFile, Map<String, ICNPTerm> terms) {
        try {
            FileInputStream file = new FileInputStream(mapFile);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            int i = 0;
            for (Row row : sheet) {
                var icnp = formatter.formatCellValue(row.getCell(0));
                var sct = formatter.formatCellValue(row.getCell(2));
                if (sct == null || sct.isEmpty() || sct.trim().isEmpty() || sct.isBlank()) {
                    sct = null;

                }
                var t = terms.get(icnp);

                if (t == null) {
                    System.out.println("What?? " + icnp);
                } else {
                    t.setSnomedTerm(sct);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();

        }
        return terms;
    }

    static Map<String, ICNPTerm> loadICNP() {
        File f = new File(ROOT, ICNP);
        if (!f.exists()) {
            throw new RuntimeException("ICNP file does not exist");
        }

        try {
            FileInputStream file = new FileInputStream(f);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            Map<String, ICNPTerm> terms = new HashMap<>();
            DataFormatter formatter = new DataFormatter();
            int i = 0;
            for (Row row : sheet) {
                var d = formatter.formatCellValue(row.getCell(1));
                var axis = formatter.formatCellValue(row.getCell(2));
                var term = formatter.formatCellValue(row.getCell(7));
                var def = formatter.formatCellValue(row.getCell(8));
                ICNPTerm t = new ICNPTerm(d, axis, term, def, null);
                terms.put(d, t);

                i++;
            }
            System.out.println("Read  " + i + " rows");
            return terms;

        } catch (IOException e) {
            throw new RuntimeException("Could not load ICNP terms ", e);
        }


    }

}
