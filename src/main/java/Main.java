import org.apache.commons.lang3.StringUtils;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.commons.*;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import java.io.*;
import java.util.*;


public class Main {

//    public static void main(String[] args) throws FileNotFoundException, UnsupportedEncodingException, ParseException {
   public static void main(String[] args) throws Exception {
        Main xread2 = new Main();
    }

    public Main() throws Exception {
      readExcel();
    }

    public static void readExcel() throws Exception {
        XSSFRow row;
        XSSFCell cell;
        String[][] value = null;
        double[][] nums = null;

        try {
//            FileInputStream inputStream = new FileInputStream("TEST.xlsx");
            FileInputStream file = new FileInputStream("/Users/matthewjames/Downloads/300000614949MeansTDSData.xlsx");
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
                                        //                                  case XSSFCell.CELL_TYPE_NUMERIC:
                                        value[r][c] = ""
                                                + cell.getNumericCellValue();
                                        break;

                                    case STRING:
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
                                System.out.print("[null]\t");
                            }
                        } // for(c)
                        System.out.print("\n");
                    }
                } // for(r)
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

//identify subentities in data model
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

// Count number of lines of sub-entities in array
        int cnt = 0;
        for (int d=0; d < value.length; d++) {
            if (subEntityName.contains(value[d][2])) {
                cnt++;
            }
        }
        System.out.println("count is " + cnt);

// Create 2D array to hold sub-entities
        String[][] subEntities = new String[cnt][value[0].length];
        if (cnt > 0) {
            int g = 0;
            for (int e = 0; e < value.length && g < cnt; e++) {
                if (subEntityName.contains(value[e][2])) {
                    subEntities[g] = value[e];
                    g++;
                    }
                }
        System.out.println(Arrays.deepToString(subEntities).replace("], ", "]\n").replace("[[", "[").replace("]]", "]"));
        }

// Create HashMap to hold relationship
        HashMap <String, List<String>> subEntityRel = new HashMap<>();
        subEntityRel.put("ADDPROPERTY", Arrays.asList(new String[]{"ADDTHIRD"}));
        subEntityRel.put("BANKACC", Arrays.asList(new String[]{"BPFAMILYEMPL"}));
        subEntityRel.put("BUSINESSPART", Arrays.asList(new String[]{"BPPROPERTY","BUSPARTBANK" }));
        subEntityRel.put("COMPANY", Arrays.asList(new String[]{"COPROPERTY","SHARE" }));
        subEntityRel.put("EMPLOYMENT_CLIENT", Arrays.asList(new String[]{"EMP_CLI_KNOWN_CHANGE","EMPLOY_BEN_CLIENT",
                "CLI_NON_HM_L17","CLI_NON_HM_WAGE_SLIP"}));
        subEntityRel.put("EMPLOYMENT_PARTNER", Arrays.asList(new String[]{"EMPLOY_BEN_PARTNER","PAR_NON_HM_L17",
                "PAR_NON_HM_WAGE_SLIP","PAR_EMPLOY_KNOWN_CHANGE" }));
        subEntityRel.put("SELFEMPLOY", Arrays.asList(new String[]{"SEFAMILYEMPL", "SELFEMPBANK", "SEPROPERTY"}));

//        System.out.println(subEntityRel);
//        List<String> subEntity = subEntityRel.get("BUSINESSPART");
//        System.out.println(subEntity);

//        System.out.println(Arrays.deepToString(value).replace("], ", "]\n").replace("[[", "[").replace("]]", "]"));
//        System.out.println(value[5][3] + "length " + value.length);
        String ref = "";

// Get all entity names in a list
        List<String> listStrings = new ArrayList<String>();
        for (int c=0; c < value.length-1; c++) {
            if (!value[c][2].equals(value[c+1][2]) && !value[c][2].equals("ENTITY_TYPE")) {
                listStrings.add(value[c][2]);
            }
        }
// Replace & with &amp;
        for (int b=0; b<listStrings.size()-1; b++) {
            String replaceLetters  = listStrings.get(b);
            replaceLetters = replaceLetters.replaceAll("&","&amp;");
            listStrings.set(b, replaceLetters);
        }
        System.out.println(listStrings);

// Write out file
        try {
            PrintWriter writer = new PrintWriter("/Users/matthewjames/Downloads/file.xml", "UTF-8");
            writer.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");   // write header
            writer.println("<session-data xmlns=\"http://oracle.com/determinations/engine/sessiondata/10.2\">"); // write xmlns data
            writer.println("<entity id =" + "\"" + "global" + "\"" + ">"); // write global entity id

            // find case reference for global instance identifier
            for (int a = 0; a < value.length - 1; a++) {
                if (value[a][2].equals("global")) {
                    ref = value[a][5];
                    //                System.out.println(ref);
                    break;
                }
            }

            writer.println("<instance id=" + "\"" + ref + "\"" + ">"); // write global instance id
            // write global records
            for (int b = 0; b < value.length; b++) {
                //            Date dateOutput = new SimpleDateFormat("yyyy-MM-dd").parse(value[b][5]);
                // Get dates in YYYY-MM-DD format
                String outYear = "";
                String outMonth = "";
                String outDay = "";
                String outDate = "";
                if (value[b][5].length() == 10) {
                    outYear = value[b][5].substring(6, 10);
                    outMonth = value[b][5].substring(3, 5);
                    outDay = value[b][5].substring(0, 2);
                    outDate = outYear + "-" + outMonth + "-" + outDay;
                    System.out.println(outDate);
                }
                if (value[b][2].equals("global")) {
                    writer.println("<attribute id=" + "\"" + value[b][4] + "\"" + ">");
                    if (value[b][5].equals("~ ~") ||
                            ((value[b][4].contains("_D_") || value[b][4].contains("DATE")) && value[b][5].contains("BLANK"))) {
                        writer.println("<uncertain-val/>");
                    } else if ((value[b][5].equals("0") || value[b][5].contains("BLANK")) && value[b][4].contains("_T_")) {
                        writer.println("<text-val/>");
                    } else if (value[b][4].contains("_T_") || value[b][4].equals("GB_INPUT_B_13WP3_49A") ||
                            value[b][4].equals("APPLICATION_CASE_REF") || value[b][4].equals("COST_LIMIT_CHANGED_FLAG")) {
                        writer.println("<text-val>" + value[b][5] + "</text-val>");
                    } else if (value[b][4].contains("_B_") || value[b][5].equals("true") || value[b][5].equals("false")) {
                        writer.println("<boolean-val>" + value[b][5] + "</boolean-val>");
                    } else if (value[b][4].contains("_D_") || value[b][4].contains("DATE")) {
                        writer.println("<date-val>" + outDate + "</date-val>");
                    } else if (value[b][4].contains("_N_")) {
                        writer.println("<number-val>" + value[b][5] + "</number-val>");
                    } else if (value[b][4].contains("_C_") || StringUtils.isNumeric(value[b][5]) || value[b][5].contains("0.00")) {
                        writer.println("<currency-val>" + value[b][5] + "</currency-val>");
                    } else if (value[b][5].contains("BLANK")) {
                        writer.println("<text-val/>");
                    } else {
                        writer.println("<text-val>" + value[b][5] + "</text-val>");
                    }
                    writer.println("</attribute>");

                }
                //            if (value[b][2].equals("OPPONENT_OTHER_PARTIES")) {
                if (listStrings.contains(value[b][2]) && !subEntityName.contains(value[b][2])) {
                    String entityName = value[b][2];
                    String inst = "";
                    inst = value[b][3];
                    if (!value[b - 1][2].equals(entityName)) {
                        writer.println("<entity id=" + "\"" + value[b][2] + "\"" + ">");
                    }
                    if (!value[b][3].equals(value[b - 1][3])) {
                        writer.println("<instance id=" + "\"" + value[b][3] + "\"" + ">");
                    }
                    writer.println("<attribute id=" + "\"" + value[b][4] + "\"" + ">");
                    if (value[b][5].equals("~ ~") ||
                            ((value[b][4].contains("_D_") || value[b][4].contains("DATE") || value[b][4].contains("DOB"))
                                    && value[b][5].contains("BLANK"))) {
                        writer.println("<uncertain-val/>");
                    } else if ((value[b][5].equals("0") || value[b][5].contains("BLANK")) && value[b][4].contains("_T_")) {
                        writer.println("<text-val/>");
                    } else if (value[b][4].contains("_T_")||value[b][4].equals("BANKACC_INPUT_N_7WP2_5A")) {
                        writer.println("<text-val>" + value[b][5] + "</text-val>");
                    } else if (entityName.equals("PROCEEDING") && value[b][4].contains("LEVEL_OF_SERVICE")) {
                        writer.println("<text-val>" + value[b][5] + "</text-val>");
                    } else if (value[b][4].contains("_B_") || value[b][5].equals("true") || value[b][5].equals("false")) {
                        writer.println("<boolean-val>" + value[b][5] + "</boolean-val>");
                    } else if (value[b][4].contains("_D_") || value[b][4].contains("DATE")) {
                        writer.println("<date-val>" + outDate + "</date-val>");
                    } else if (value[b][4].contains("_N_")) {
                        writer.println("<number-val>" + value[b][5] + "</number-val>");
                    } else if (value[b][4].contains("_C_") || StringUtils.isNumeric(value[b][5]) ||
                            value[b][5].contains("0.00")) {
                        writer.println("<currency-val>" + value[b][5] + "</currency-val>");
                    } else if (value[b][5].contains("BLANK")) {
                        writer.println("<text-val/>");
                    } else {
                        writer.println("<text-val>" + value[b][5] + "</text-val>");
                    }
                    writer.println("</attribute>");
// add any sub-entities using subEntities array
                   if ((cnt > 0 && subEntityRel.containsKey(value[b][2]) && !value[b][3].equals(value[b + 1][3]))) {
                        for (int h = 0; h < subEntities.length; h++) {
                            if (subEntityRel.get(value[b][2]).contains(subEntities[h][2])) {
                                String outDate2 = "";
                                if (subEntities[h][5].length() == 10) {
                                    outDate2 = subEntities[h][5].substring(6, 10) + "-" + subEntities[h][5].substring(3, 5)
                                            + "-" + subEntities[h][5].substring(0, 2);
                                }
                                if (h == 0 || !subEntities[h - 1][2].equals(subEntities[h][2])) {
                                    writer.println("<entity id=" + "\"" + subEntities[h][2] + "\"" + ">");
                                }
                                if (h == 0 || !subEntities[h][3].equals(subEntities[h - 1][3])) {
                                    writer.println("<instance id=" + "\"" + subEntities[h][3] + "\"" + ">");
                                }
                                writer.println("<attribute id=" + "\"" + subEntities[h][4] + "\"" + ">");
                                if (subEntities[h][5].equals("~ ~") ||
                                        ((subEntities[h][4].contains("_D_") || subEntities[h][4].contains("DATE")) && subEntities[h][5].contains("BLANK"))) {
                                    writer.println("<uncertain-val/>");
                                } else if ((subEntities[h][5].equals("0") || subEntities[h][5].contains("BLANK")) && subEntities[h][4].contains("_T_")) {
                                    writer.println("<text-val/>");
                                } else if (subEntities[h][4].contains("_T_")) {
                                    writer.println("<text-val>" + subEntities[h][5] + "</text-val>");
                                } else if (subEntities[h][4].contains("_B_") || subEntities[h][5].equals("true") || subEntities[h][5].equals("false")) {
                                    writer.println("<boolean-val>" + subEntities[h][5] + "</boolean-val>");
                                } else if (subEntities[h][4].contains("_D_") || subEntities[h][4].contains("DATE")) {
                                    writer.println("<date-val>" + outDate2 + "</date-val>");
                                } else if (subEntities[h][4].contains("_N_")) {
                                    writer.println("<number-val>" + subEntities[h][5] + "</number-val>");
                                } else if (subEntities[h][4].contains("_C_") || StringUtils.isNumeric(subEntities[h][5]) || subEntities[h][5].contains("0.00")) {
                                    writer.println("<currency-val>" + subEntities[h][5] + "</currency-val>");
                                } else if (subEntities[h][5].contains("BLANK")) {
                                    writer.println("<text-val/>");
                                } else {
                                    writer.println("<text-val>" + subEntities[h][5] + "</text-val>");
                                }
                                writer.println("</attribute>");
                                if (cnt == 1 || (!value[h][3].equals(value[h + 1][3]))) {
                                    writer.println("</instance>");
                                }
                                if (cnt == 1 || !value[h + 1][2].equals(entityName)) {
                                    writer.println("</entity>");
                                }
                            }
                        }
                    }
                    if (!value[b][3].equals(value[b + 1][3])) {
                        writer.println("</instance>");
                    }
                    if (!value[b + 1][2].equals(entityName)) {
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

        System.out.println(ref);
    }
}
