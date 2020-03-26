import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.*;
import java.util.*;


public class Main {

    private static String inDate;

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

// Set path of output file

    String MeansXDSFile = "/Users/matthewjames/Downloads/file.xds";

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

                            } else {
//                                System.out.print("[null]\t");
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
//
//identify subentities in data model
    static ArrayList<String> SubEntityName () {
        ArrayList<String> subEntityName = new ArrayList<String>();
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

    {
//        System.out.println(SubEntityName());

// Call readExcel file method

        String[][] meansData = readExcel(MeansXLFile);

// Get number of rows and columns
        int meansFileLength = meansData.length;
        int meansFileCol = meansData[0].length;

//        System.out.println("means file length " + meansFileLength + "cols " + meansFileCol);
//        System.out.println("data in row 1 col 1" + meansData[1][8]);


// Count number of lines of sub-entities in array
        int cnt = 0;
        for (int d = 0; d < meansFileLength; d++) {
            if (SubEntityName().contains(meansData[d][2])) {
                cnt++;
            }
        }
//        System.out.println("count is " + cnt);

// Create 2D array to hold sub-entities
        String[][] subEntities = new String[cnt][meansFileCol];
        if (cnt > 0) {
            int g = 0;
            for (int e = 0; e < meansFileLength && g < cnt; e++) {
                if (SubEntityName().contains(meansData[e][2])) {
                    subEntities[g] = meansData[e];
                    g++;
                }
            }
//            System.out.println(Arrays.deepToString(subEntities).replace("], ", "]\n").replace("[[", "[").replace("]]", "]"));
        }

// Create HashMap to hold relationship
        HashMap<String, List<String>> subEntityRel = new HashMap<>();
        subEntityRel.put("ADDPROPERTY", Arrays.asList(new String[]{"ADDTHIRD"}));
        subEntityRel.put("BANKACC", Arrays.asList(new String[]{"BPFAMILYEMPL"}));
        subEntityRel.put("BUSINESSPART", Arrays.asList(new String[]{"BPPROPERTY", "BUSPARTBANK"}));
        subEntityRel.put("COMPANY", Arrays.asList(new String[]{"COPROPERTY", "SHARE"}));
        subEntityRel.put("EMPLOYMENT_CLIENT", Arrays.asList(new String[]{"EMP_CLI_KNOWN_CHANGE", "EMPLOY_BEN_CLIENT",
                "CLI_NON_HM_L17", "CLI_NON_HM_WAGE_SLIP"}));
        subEntityRel.put("EMPLOYMENT_PARTNER", Arrays.asList(new String[]{"EMPLOY_BEN_PARTNER", "PAR_NON_HM_L17",
                "PAR_NON_HM_WAGE_SLIP", "PAR_EMPLOY_KNOWN_CHANGE"}));
        subEntityRel.put("SELFEMPLOY", Arrays.asList(new String[]{"SEFAMILYEMPL", "SELFEMPBANK", "SEPROPERTY"}));

//        System.out.println(subEntityRel);

        String ref = "";

// Get all entity names in a list
        List<String> listStrings = new ArrayList<String>();
        for (int c = 0; c < meansFileLength-1; c++) {
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
//        System.out.println("listStrings is " + listStrings);

// Write out converted file
        try {
            PrintWriter writer = new PrintWriter(MeansXDSFile, "UTF-8");
            writer.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");   // write header
            writer.println("<session-data xmlns=\"http://oracle.com/determinations/engine/sessiondata/10.2\">"); // write xmlns data
            writer.println("<entity id =" + "\"" + "global" + "\"" + ">"); // write global entity id

            // find case reference for global instance identifier
            for (int a = 0; a < meansFileLength - 1; a++) {
                if (meansData[a][2].equals("global")) {
                    ref = meansData[a][5];
                    //                System.out.println(ref);
                    break;
                }
            }

            writer.println("<instance id=" + "\"" + ref + "\"" + ">"); // write global instance id
            // write global records
            for (int b = 1; b < meansFileLength-1; b++) {
//write out global records first i.e. lowest level
                String entityName = meansData[b][2];
                String entPrevLn = meansData[b - 1][2];
                String inst = meansData[b][3];
                String instPrevLn = meansData[b - 1][3];
                String attrName = meansData[b][4];
                String attrVal = meansData[b][5];
                String instNextLn = meansData[b + 1][3];
                String entNextLn = meansData[b + 1][2];
//                System.out.println(entityName + entPrevLn + inst + instPrevLn + instNextLn + attrName + attrVal);
                if (entityName.equals("global")) {
                    writer.println("<attribute id=" + "\"" + attrName + "\"" + ">");
                    writer.println(xdsDataStructure(meansData, b));
/*                    if ((attrVal.equals("0") || attrVal.contains("BLANK")) && attrName.contains("_T_")) {
                        writer.println("<text-val/>");
                    } else if (attrVal.contains("~ ~") ||
                            ((attrName.contains("_D_") || attrName.contains("DATE")) && attrVal.contains("BLANK"))) {
                        writer.println("<uncertain-val/>");
                    } else if ((attrVal.equals("0") || attrVal.contains("BLANK")) && attrName.contains("_T_")) {
                        writer.println("<text-val/>");
                    } else if (attrName.contains("_T_") || attrName.equals("GB_INPUT_B_13WP3_49A") ||
                            attrName.equals("APPLICATION_CASE_REF") || attrName.equals("COST_LIMIT_CHANGED_FLAG")) {
                        writer.println("<text-val>" + attrVal + "</text-val>");
                    } else if (attrName.contains("_B_") || attrVal.equals("true") || attrVal.equals("false")) {
                        writer.println("<boolean-val>" + attrVal + "</boolean-val>");
                    } else if (attrName.contains("_D_") || attrName.contains("DATE")) {
                        writer.println("<date-val>" + yrFirstDate(attrVal) + "</date-val>");
                    } else if (attrName.contains("_N_")) {
                        writer.println("<number-val>" + attrVal + "</number-val>");
                    } else if (attrName.contains("_C_") || StringUtils.isNumeric(attrVal) || attrVal.contains("0.00")) {
                        writer.println("<currency-val>" + attrVal + "</currency-val>");
                    } else if (attrVal.contains("BLANK")) {
                        writer.println("<text-val/>");
                    } else {
                        writer.println("<text-val>" + attrVal + "</text-val>");
                    }
*/                    writer.println("</attribute>");
                }
// write out entity level records
                if (listStrings.contains(entityName) && !SubEntityName().contains(entityName)) {
                    if (!entPrevLn.equals(entityName)) {
                        writer.println("<entity id=" + "\"" + entityName + "\"" + ">");
                    }
                    if (!inst.equals(instPrevLn)) {
                        writer.println("<instance id=" + "\"" + inst + "\"" + ">");
                    }
                    writer.println("<attribute id=" + "\"" + attrName + "\"" + ">");
                    writer.println(xdsDataStructure(meansData, b));
/*                    if (attrVal.equals("~ ~") ||
                            ((attrName.contains("_D_") || attrName.contains("DATE") || attrName.contains("DOB"))
                                    && attrVal.contains("BLANK"))) {
                        writer.println("<uncertain-val/>");
                    } else if ((attrVal.equals("0") || attrVal.contains("BLANK")) && attrName.contains("_T_")) {
                        writer.println("<text-val/>");
                    } else if (attrName.contains("_T_") || attrName.equals("BANKACC_INPUT_N_7WP2_5A")) {
                        writer.println("<text-val>" + attrVal + "</text-val>");
                    } else if (entityName.equals("PROCEEDING") && attrName.contains("LEVEL_OF_SERVICE")) {
                        writer.println("<text-val>" + attrVal + "</text-val>");
                    } else if (attrName.contains("_B_") || attrVal.equals("true") || attrVal.equals("false")) {
                        writer.println("<boolean-val>" + attrVal + "</boolean-val>");
                    } else if (attrName.contains("_D_") || attrName.contains("DATE")) {
                        writer.println("<date-val>" + yrFirstDate(attrVal) + "</date-val>");
                    } else if (attrName.contains("_N_")) {
                        writer.println("<number-val>" + attrVal + "</number-val>");
                    } else if (attrName.contains("_C_") || StringUtils.isNumeric(attrVal) ||
                            attrVal.contains("0.00")) {
                        writer.println("<currency-val>" + attrVal + "</currency-val>");
                    } else if (attrVal.contains("BLANK")) {
                        writer.println("<text-val/>");
                    } else {
                        writer.println("<text-val>" + attrVal + "</text-val>");
                    }*/
                    writer.println("</attribute>");
    // add any sub-entities using subEntities array
                    if ((cnt > 0 && subEntityRel.containsKey(entityName) && !inst.equals(instNextLn))) {
                        for (int h = 0; h < subEntities.length; h++) {
                            if (subEntityRel.get(entityName).contains(subEntities[h][2])) {
                                if (h == 0 || !subEntities[h - 1][2].equals(subEntities[h][2])) {
                                    writer.println("<entity id=" + "\"" + subEntities[h][2] + "\"" + ">");
                                }
                                if (h == 0 || !subEntities[h][3].equals(subEntities[h - 1][3])) {
                                    writer.println("<instance id=" + "\"" + subEntities[h][3] + "\"" + ">");
                                }
                                writer.println("<attribute id=" + "\"" + subEntities[h][4] + "\"" + ">");
                                writer.println(xdsDataStructure(subEntities, h));
/*                                if (subEntities[h][5].equals("~ ~") ||
                                        ((subEntities[h][4].contains("_D_") || subEntities[h][4].contains("DATE")) && subEntities[h][5].contains("BLANK"))) {
                                    writer.println("<uncertain-val/>");
                                } else if ((subEntities[h][5].equals("0") || subEntities[h][5].contains("BLANK")) && subEntities[h][4].contains("_T_")) {
                                    writer.println("<text-val/>");
                                } else if (subEntities[h][4].contains("_T_")) {
                                    writer.println("<text-val>" + subEntities[h][5] + "</text-val>");
                                } else if (subEntities[h][4].contains("_B_") || subEntities[h][5].equals("true") || subEntities[h][5].equals("false")) {
                                    writer.println("<boolean-val>" + subEntities[h][5] + "</boolean-val>");
                                } else if (subEntities[h][4].contains("_D_") || subEntities[h][4].contains("DATE")) {
                                    writer.println("<date-val>" + yrFirstDate(subEntities[h][5]) + "</date-val>");
                                } else if (subEntities[h][4].contains("_N_")) {
                                    writer.println("<number-val>" + subEntities[h][5] + "</number-val>");
                                } else if (subEntities[h][4].contains("_C_") || StringUtils.isNumeric(subEntities[h][5]) || subEntities[h][5].contains("0.00")) {
                                    writer.println("<currency-val>" + subEntities[h][5] + "</currency-val>");
                                } else if (subEntities[h][5].contains("BLANK")) {
                                    writer.println("<text-val/>");
                                } else {
                                    writer.println("<text-val>" + subEntities[h][5] + "</text-val>");
                                } */
                                writer.println("</attribute>");
                                if (cnt == 1 || (!subEntities[h][3].equals(subEntities[h + 1][3]))) {
                                    writer.println("</instance>");
                                }
                                if (cnt == 1 || !subEntities[h + 1][2].equals(entityName)) {
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
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
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

    public static String xdsDataStructure(String [][] meansData, int b) {
        String attrName = meansData[b][4];
        String attrVal = meansData[b][5];
        if ((attrVal.equals("0") || attrVal.contains("BLANK")) && attrName.contains("_T_")) {
            return ("<text-val/>");
        } else if (attrVal.contains("~ ~") ||
                ((attrName.contains("_D_") || attrName.contains("DATE")) && attrVal.contains("BLANK"))) {
            return ("<uncertain-val/>");
        } else if ((attrVal.equals("0") || attrVal.contains("BLANK")) && attrName.contains("_T_")) {
            return ("<text-val/>");
        } else if (attrName.contains("_T_") || attrName.equals("GB_INPUT_B_13WP3_49A") ||
                attrName.equals("APPLICATION_CASE_REF") || attrName.equals("COST_LIMIT_CHANGED_FLAG")) {
            return ("<text-val>" + attrVal + "</text-val>");
        } else if (attrName.contains("_B_") || attrVal.equals("true") || attrVal.equals("false")) {
            return ("<boolean-val>" + attrVal + "</boolean-val>");
        } else if (attrName.contains("_D_") || attrName.contains("DATE")) {
            return ("<date-val>" + yrFirstDate(attrVal) + "</date-val>");
        } else if (attrName.contains("_N_")) {
            return("<number-val>" + attrVal + "</number-val>");
        } else if (attrName.contains("_C_") || StringUtils.isNumeric(attrVal) || attrVal.contains("0.00")) {
            return ("<currency-val>" + attrVal + "</currency-val>");
        } else if (attrVal.contains("BLANK")) {
            return ("<text-val/>");
        } else {
            return ("<text-val>" + attrVal + "</text-val>");
        }
//        return("</attribute>");
    }

}
