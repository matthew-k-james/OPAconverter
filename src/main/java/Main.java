import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;


public class Main {

    public Main() throws Exception {
    }

    public static void main(String[] args) throws Exception {
        Main xread = new Main();
    }

//* Assumptions: Excel spreadsheet consists of 7 columns. Only 4 are used:
//      Column 3 (becomes 2 when file is read) : ENTITY_TYPE - contains entity name. Global is 'global'.
//      Column 4 (becomes 3 when file is read) : ENTITY_ID - contains the instance ID.
//      Column 5 (becomes 4 when file is read) : ATTRIBUTE_NAME - contains the attribute ID.
//      Column 6 (becomes 5 when file is read) : ATTRIBUTE_VALUE - contains the value. '~ ~' signifies uncertain.
//                  0 value in a text field signifies a blank. Date values are 'DD-MM-YYYY'. '&' values aren't escaped.
//
// Set path of Excel File containing Means data

    String MeansXLFile = "/Users/matthewjames/Downloads/300000614949MeansTDSData.xlsx";

    String MeansDataModel = "/Users/matthewjames/Downloads/Means 18A Data Model.xlsx";

// Set path of output file

    String MeansXDSFile = "/Users/matthewjames/Downloads/MeansXDSFile2.xds";

    public static String[][] readExcel(String XLFile) throws Exception {
        XSSFRow row;
        XSSFCell cell;
        String[][] value = null;
//        double[][] nums = null;

        if (XLFile == null) {
            throw new IllegalArgumentException("the means Excel file name must be provided");
        }

        try {
            FileInputStream file = new FileInputStream(XLFile);
//            FileInputStream file = new FileInputStream("/Users/matthewjames/Downloads/300000614949MeansTDSData.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // get sheet number
            int sheetCn = workbook.getNumberOfSheets();

            for (int cn = 0; cn < sheetCn; cn++) {

                // get 0th sheet data
                XSSFSheet sheet = workbook.getSheetAt(cn);

                // get number of rows from sheet
                int rows = sheet.getPhysicalNumberOfRows();

                // get number of cell from row
                int cells = sheet.getRow(cn).getPhysicalNumberOfCells();

                value = new String[rows][cells];

                for (int r = 0; r < rows; r++) {
                    row = sheet.getRow(r); // bring row
                    if (row != null) {
                        for (int c = 0; c < cells; c++) {
                            cell = row.getCell(c);
//                            nums = new double[rows][cells];

                            if (cell != null) {

                                switch (cell.getCellType()) {

                                    case NUMERIC:
// We aren't interested in the numeric cells so we ignore them
//                                  case XSSFCell.CELL_TYPE_NUMERIC:
//                                        value[r][c] = ""
//                                                + cell.getNumericCellValue();
                                        break;

                                    case STRING:
// Read in Strings and escape any '&' chars in excel spreadsheet
//                                    case XSSFCell.CELL_TYPE_STRING:
                                        value[r][c] = ""
                                                + cell.getStringCellValue().replace("&", "&amp;");
                                        break;

                                    case BLANK:
//                                    case XSSFCell.CELL_TYPE_BLANK:
                                        value[r][c] = "[BLANK]";
                                        break;

                                    default:
                                }
//                                System.out.print(value[r][c]);

                            }
                        } // for(c)
//                        System.out.print("\n");
                    }
                } // for(r)
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return value;
    }

/*    {
        for (int z=0; z < 10; z++) {
            System.out.println(readExcel(MeansDataModel)[z][1] + " " + readExcel(MeansDataModel)[z][2] + " " +
                    readExcel(MeansDataModel)[z][3] + " " + readExcel(MeansDataModel)[z][4] + " " + readExcel(MeansDataModel)[z][5]);
        }
    }*/
// Call readExcel file method for source file

    String[][] meansData = readExcel(MeansXLFile);
    String[][] dataModel = readExcel(MeansDataModel);

    public static HashMap<String, String> readMeansDataModel(String[][] dataModel) throws Exception {
        HashMap<String, String> modelMap = new HashMap<>();
        for (int l=1; l < dataModel.length; l++) {
            modelMap.put(dataModel[l][4], dataModel[l][2]);
        }
//        modelMap.entrySet().forEach(entry->{
//            System.out.println(entry.getKey() + " " + entry.getValue());
//        });
        return modelMap;
    }
    //
//identify subentities in data model
    private static ArrayList<String> subEntityName() {
        ArrayList<String> subEntityName = new ArrayList<>();
        subEntityName.add("ADDTHIRD");
        subEntityName.add("BPFAMILYEMPL");
        subEntityName.add("BPPROPERTY");
        subEntityName.add("BUSPARTBANK");
        subEntityName.add("COPROPERTY");
        subEntityName.add("SHARE");
        subEntityName.add("EMP_CLI_KNOWN_CHANGE");
        subEntityName.add("EMPLOY_BEN_CLIENT");
        subEntityName.add("CLI_NON_HM_L17");
        subEntityName.add("CLI_NON_HM_WAGE_SLIP");
        subEntityName.add("EMPLOY_BEN_PARTNER");
        subEntityName.add("PAR_NON_HM_L17");
        subEntityName.add("PAR_NON_HM_WAGE_SLIP");
        subEntityName.add("PAR_EMPLOY_KNOWN_CHANGE");
        subEntityName.add("SEFAMILYEMPL");
        subEntityName.add("SELFEMPBANK");
        subEntityName.add("SEPROPERTY");
        return subEntityName;
    }



// Get number of rows and columns
    public int getNumRows(String[][] meansData){
        return meansData.length;
    }

    public int getNumCols(String[][] meansData){
        return meansData[0].length;
    }


    private int getNumSubEnts(String[][] meansData) {
// Count number of lines of sub-entities in array
        int cnt = 0;
        for (int d = 0; d < getNumRows(meansData); d++) {
            if (subEntityName().contains(meansData[d][2])) {
                cnt++;
            }
        }
        return cnt;
    }

// Create 2D array to hold sub-entities
    private String[][] getSubEntities(String[][] meansData){
        int count = getNumSubEnts(meansData);
        String[][] subEntities = new String[count][getNumCols(meansData)];
        if (count > 0) {
            int g = 0;
            for (int e = 0; e < getNumRows(meansData) && g < count; e++) {
                if (subEntityName().contains(meansData[e][2])) {
                    subEntities[g] = meansData[e];
                    g++;
                }
            }
//            System.out.println(Arrays.deepToString(subEntities).replace("], ", "]\n").replace("[[", "[").replace("]]", "]"));
        }
        return subEntities;
    }


/*    {
        int xx = 0;
        for (xx=0; xx < getNumSubEnts(meansData); xx++) {
            System.out.println(getSubEntities(meansData)[xx][2]);
        }
    }
*/
// Create HashMap to hold relationship
    public HashMap<String, List<String>> subEntRel() {
        HashMap<String, List<String>> subEntityRel = new HashMap<>();
        subEntityRel.put("ADDPROPERTY", Arrays.asList("ADDTHIRD"));
        subEntityRel.put("BANKACC", Arrays.asList("BPFAMILYEMPL"));
        subEntityRel.put("BUSINESSPART", Arrays.asList("BPPROPERTY", "BUSPARTBANK"));
        subEntityRel.put("COMPANY", Arrays.asList("COPROPERTY", "SHARE"));
        subEntityRel.put("EMPLOYMENT_CLIENT", Arrays.asList("EMP_CLI_KNOWN_CHANGE", "EMPLOY_BEN_CLIENT",
                "CLI_NON_HM_L17", "CLI_NON_HM_WAGE_SLIP"));
        subEntityRel.put("EMPLOYMENT_PARTNER", Arrays.asList("EMPLOY_BEN_PARTNER", "PAR_NON_HM_L17",
                "PAR_NON_HM_WAGE_SLIP", "PAR_EMPLOY_KNOWN_CHANGE"));
        subEntityRel.put("SELFEMPLOY", Arrays.asList("SEFAMILYEMPL", "SELFEMPBANK", "SEPROPERTY"));
        return subEntityRel;
    }
//
//{
//    System.out.println(subEntRel());
//}

//        String ref = "";

// Get all entity names in a list
    private List<String> getEntityList(String[][] meansData) {
        List<String> listStrings = new ArrayList<String>();
        for (int c = 0; c < getNumRows(meansData) - 1; c++) {
            if (!meansData[c][2].equals(meansData[c + 1][2]) && !meansData[c][2].equals("ENTITY_TYPE")) {
//                System.out.println(meansData[c][2]);
                listStrings.add(meansData[c][2]);
            }
// Replace & with &amp;
            for (int b = 0; b < listStrings.size() - 1; b++) {
                String replaceLetters = listStrings.get(b);
                replaceLetters = replaceLetters.replaceAll("&", "&amp;");
                listStrings.set(b, replaceLetters);
            }
        }
        return listStrings;
    }

//{
//    System.out.println(getEntityList(meansData));
//}
// Write out converted file
    {
//    public void writeXDSFile(String[][] meansData) throws FileNotFoundException, UnsupportedEncodingException {
            PrintWriter writer = new PrintWriter(MeansXDSFile, StandardCharsets.UTF_8);
            String ref = "";

            writer.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");   // write header
            writer.println("<session-data xmlns=\"http://oracle.com/determinations/engine/sessiondata/10.2\">"); // write xmlns data
            writer.println("<entity id =" + "\"" + "global" + "\"" + ">"); // write global entity id

            // find case reference for global instance identifier
            for (int a = 0; a < getNumRows(meansData) - 1; a++) {
                if (meansData[a][2].equals("global")) {
                    ref = meansData[a][5];
                    System.out.println(ref);
                    break;
                }
            }

            writer.println("<instance id=" + "\"" + ref + "\"" + ">"); // write global instance id
            // write global records
            for (int b = 1; b < getNumRows(meansData) - 1; b++) {
//write out global records first i.e. lowest level
                String entityName = meansData[b][2];
                String entPrevLn = meansData[b - 1][2];
                String inst = meansData[b][3];
                String instPrevLn = meansData[b - 1][3];
                String attrName = meansData[b][4];
                String instNextLn = meansData[b + 1][3];
                String entNextLn = meansData[b + 1][2];
                if (entityName.equals("global")) {
                    writer.println("<attribute id=" + "\"" + attrName + "\"" + ">");
                    writer.println(xdsDataStructure(dataModel, meansData, b));
                    writer.println("</attribute>");
                }
// write out entity level records
                if (getEntityList(meansData).contains(entityName) && !subEntityName().contains(entityName)) {
                    if (!entPrevLn.equals(entityName)) {
                        writer.println("<entity id=" + "\"" + entityName + "\"" + ">");
                    }
                    if (!inst.equals(instPrevLn)) {
                        writer.println("<instance id=" + "\"" + inst + "\"" + ">");
                    }
                    writer.println("<attribute id=" + "\"" + attrName + "\"" + ">");
                    writer.println(xdsDataStructure(dataModel, meansData, b));
                    writer.println("</attribute>");
// add any sub-entities using subEntities array
                    int count = getNumSubEnts(meansData);
                    if ((count > 0 && subEntRel().containsKey(entityName) && !inst.equals(instNextLn))) {
                        for (int h = 0; h < getSubEntities(meansData).length; h++) {
                            if (subEntRel().get(entityName).contains(getSubEntities(meansData)[h][2])) {
                                if (h == 0 || !getSubEntities(meansData)[h - 1][2].equals(getSubEntities(meansData)[h][2])) {
                                    writer.println("<entity id=" + "\"" + getSubEntities(meansData)[h][2] + "\"" + ">");
                                }
                                if (h == 0 || !getSubEntities(meansData)[h][3].equals(getSubEntities(meansData)[h - 1][3])) {
                                    writer.println("<instance id=" + "\"" + getSubEntities(meansData)[h][3] + "\"" + ">");
                                }
                                writer.println("<attribute id=" + "\"" + getSubEntities(meansData)[h][4] + "\"" + ">");
                                writer.println(xdsDataStructure(getSubEntities(meansData), h));
                                writer.println("</attribute>");
                                if (count == 1 || (!getSubEntities(meansData)[h][3].equals(getSubEntities(meansData)[h + 1][3]))) {
                                    writer.println("</instance>");
                                }
                                if (count == 1 || !getSubEntities(meansData)[h + 1][2].equals(entityName)) {
                                    writer.println("</entity>");
                                }
                            }
                        }
                    }
                    if (!inst.equals(instNextLn)) {
                        writer.println("</instance>");
                    }
                    if (!entNextLn.equals(entityName)) {
                        writer.println("</entity>");
                    }
                }
            }
            writer.println("</instance>");
            writer.println("</entity>");
            writer.println("</session-data>");
            writer.close();
        }



    // Method to output date in YYYY-MM-DD format where input date is MM-YY-YYYY
//
    public static String yrFirstDate(String inDate) {
        String outYear = "";
        String outMonth = "";
        String outDay = "";
        String outDate = "";
        if (inDate.length() == 10) {
            outYear = inDate.substring(6, 10);
            outMonth = inDate.substring(3, 5);
            outDay = inDate.substring(0, 2);
            outDate = outYear + "-" + outMonth + "-" + outDay;
            return outDate;
        }
        return "";
    }

    public static String xdsDataStructure(String[][] meansData, int b) {
        String attrName = meansData[b][4];
        String attrVal = meansData[b][5];
        if ((attrVal.equals("0") || attrVal.contains("BLANK")) && attrName.contains("_T_")) {
            return ("<text-val/>");
        } else if (attrVal.contains("~ ~") ||
                ((attrName.contains("_D_") || attrName.contains("DATE")) && attrVal.contains("BLANK"))) {
            return ("<uncertain-val/>");
        } else if (attrName.contains("_T_") || attrName.equals("GB_INPUT_B_13WP3_49A") ||
                attrName.equals("APPLICATION_CASE_REF") || attrName.equals("COST_LIMIT_CHANGED_FLAG")) {
            return ("<text-val>" + attrVal + "</text-val>");
        } else if (attrName.contains("_B_") || attrVal.equals("true") || attrVal.equals("false")) {
            return ("<boolean-val>" + attrVal + "</boolean-val>");
        } else if (attrName.contains("_D_") || attrName.contains("DATE")) {
            return ("<date-val>" + yrFirstDate(attrVal) + "</date-val>");
        } else if (attrName.contains("_N_")) {
            return ("<number-val>" + attrVal + "</number-val>");
        } else if (attrName.contains("_C_") || StringUtils.isNumeric(attrVal) || attrVal.contains("0.00")) {
            return ("<currency-val>" + attrVal + "</currency-val>");
        } else if (attrVal.contains("BLANK")) {
            return ("<text-val/>");
        } else {
            return ("<text-val>" + attrVal + "</text-val>");
        }
//        return("</attribute>");
    }

    public static String xdsDataStructure(String[][] dataModel, String[][] meansData, int b) throws Exception {
        String attrName = meansData[b][4];
//        System.out.print("attribute name is " + attrName + " for line " + b);
        String attrVal = meansData[b][5];
//        System.out.println(" value is " + attrVal);
        String attrDataType = readMeansDataModel(dataModel).get(attrName);
        if (attrDataType == null) {
            System.out.println("ERROR - attribute " + attrName + " has no entry in the Data Model!!!");
            return ("ERROR!! Null value!!");
        }
        if ((attrVal.equals("0") || attrVal.contains("BLANK")) && attrDataType.contains("Text")) {
            return ("<text-val/>");
        } else if (attrVal.contains("~ ~") ||
                (attrDataType.contains("Date") && attrVal.contains("BLANK"))) {
            return ("<uncertain-val/>");
        } else if (attrDataType.contains("Text")) {
            return ("<text-val>" + attrVal + "</text-val>");
        } else if (attrDataType.contains("Boolean") || attrVal.equals("true") || attrVal.equals("false")) {
            return ("<boolean-val>" + attrVal + "</boolean-val>");
        } else if (attrDataType.contains("Date")) {
            return ("<date-val>" + yrFirstDate(attrVal) + "</date-val>");
        } else if (attrDataType.contains("Number")) {
            return ("<number-val>" + attrVal + "</number-val>");
        } else if (attrDataType.contains("Currency") || StringUtils.isNumeric(attrVal) || attrVal.contains("0.00")) {
            return ("<currency-val>" + attrVal + "</currency-val>");
        } else if (attrVal.contains("BLANK")) {
            return ("<text-val/>");
        } else {
            return ("<text-val>" + attrVal + "</text-val>");
        }
    }

}