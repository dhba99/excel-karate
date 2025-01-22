package util;

import com.intuit.karate.graal.JsMap;
import net.minidev.json.JSONArray;
import net.minidev.json.JSONObject;
import net.minidev.json.JSONValue;
import net.minidev.json.parser.JSONParser;
import net.minidev.json.parser.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.Map;

public class ExcelUtil {

    public static JSONObject toJson(String path) throws IOException {
        FileInputStream file = new FileInputStream(new File(path));

        //Crea instancia de archivo excel
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        //Obtiene la primera hoja de calculo
        XSSFSheet sheet = workbook.getSheetAt(0);

        ArrayList<String> headers = new ArrayList<>();
        ArrayList<ArrayList<String>> data = new ArrayList<>();


        //Iterate through each rows one by one
        Iterator<Row> rowIterator = sheet.iterator();
        Row row;

        if(rowIterator.hasNext()){
            row = rowIterator.next();

            Iterator<Cell> cellIterator = row.cellIterator();
            Cell cell;

            while(cellIterator.hasNext()){
                cell = cellIterator.next();

                switch (cell.getCellType()){
                    case NUMERIC:
                        headers.add(String.valueOf(cell.getNumericCellValue()));
                        break;
                    case STRING:
                        headers.add(cell.getStringCellValue());
                        break;
                    default:
                        headers.add("");
                        break;
                }
            }
        }


        while (rowIterator.hasNext()) {

            row = rowIterator.next();

            Iterator<Cell> cellIterator = row.cellIterator();
            ArrayList<String> rowData = new ArrayList<>();
            for(int i=0;i<headers.size();++i){
                Cell cell = row.getCell(i);

                switch (cell.getCellType()){
                    case NUMERIC:
                        rowData.add(String.valueOf(cell.getNumericCellValue()));
                        break;
                    case STRING:
                        rowData.add(cell.getStringCellValue());
                        break;
                    default:
                        rowData.add("");
                        break;
                }
            }
            data.add(rowData);
        }

        JSONObject jsonObject = new JSONObject();
        JSONArray jsonArrayHeaders = new JSONArray();
        JSONArray jsonArrayData = new JSONArray();

        for(int i=0;i<headers.size();++i){
            jsonArrayHeaders.add(headers.get(i));
        }

        for(int rowIt=0;rowIt<data.size();++rowIt){
            JSONObject rowObj = new JSONObject();
            for(int colIt=0;colIt<headers.size();++colIt){
                rowObj.put(headers.get(colIt), data.get(rowIt).get(colIt));
            }
            jsonArrayData.add(rowObj);
        }

        jsonObject.put("headers", jsonArrayHeaders);
        jsonObject.put("data", jsonArrayData);

        file.close();

        return jsonObject;
    }

    public static void toExcel(String json){

//        System.out.println((Polyg)jsonMap.get("headers"));
//        System.out.println("acacac");
//        Object[] arrayHeaders =  (Object[]) (Object) jsonMap.get("headers");
//        Object[] arrayData =  (Object[]) jsonMap.get("data");
        JSONObject obj ;

        // Parsear la cadena JSON a JSONObject
        JSONParser parser = new JSONParser();
        try {
            obj = (JSONObject) parser.parse(json);


        } catch (ParseException e) {
            e.printStackTrace();
            return;
        }

        JSONArray jsonArrayHeaders = (JSONArray) obj.get("headers");
        JSONArray jsonArrayData = (JSONArray) obj.get("data");

        ArrayList<String> headers = new ArrayList<>();
//        Collections.addAll(headers, arrayHeaders);
        ArrayList<ArrayList<String>> data = new ArrayList<>();

        //Obtiene los valores del JSON
        for(Object value : jsonArrayHeaders){
            headers.add(value.toString());
        }
        for(Object jsonRowData : jsonArrayData){
            JSONObject jsonObject = (JSONObject) jsonRowData;
            ArrayList<String> rowData = new ArrayList<>();
            for(int colIt=0;colIt<headers.size();++colIt){
                Object cellValue = jsonObject.get(headers.get(colIt));
                rowData.add(cellValue.toString());
//                if(cellValue instanceof String)
//                    rowData.add((String) cellValue);
//                else if(cellValue instanceof Integer)
//                    rowData.add("(Integer) cellValue");
//                else
//                    rowData.add("");
            }
            data.add(rowData);
        }

        //Crea un archivo excel a partir de los valores obtenidos

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("data");

        Row rowHeaders = sheet.createRow(0);
        for(int colIt=0;colIt<headers.size();++colIt){
            Cell cell = rowHeaders.createCell(colIt);
            cell.setCellValue(headers.get(colIt));
        }

        for(int rowIt=0;rowIt<data.size();++rowIt){
            Row row = sheet.createRow(rowIt+1);
            for(int colIt=0;colIt<headers.size();++colIt){
                Cell cell = row.createCell(colIt);
                cell.setCellValue(data.get(rowIt).get(colIt));
//                if(obj instanceof String)
//                    cell.setCellValue((String)obj);
//                else if(obj instanceof Integer)
//                    cell.setCellValue((Integer)obj);
            }
        }



        try {
            FileOutputStream out = new FileOutputStream(new File("excel_test.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }
}
