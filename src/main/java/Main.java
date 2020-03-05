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

/*identify subentities in data model
        Set<String> subEntityNames = new HashSet<String>();
        subEntityNames.add("ADDTHIRD");
        subEntityNames.add("BPFAMILYEMPL");
        subEntityNames.add("BPPROPERTY");
        subEntityNames.add("BUSPARTBANK");
        subEntityNames.add("COPROPERTY");
        subEntityNames.add("SHARE");
        subEntityNames.add("EMP_CLI_KNOWN_CHANGE");
        subEntityNames.add("EMPLOY_BEN_CLIENT");
        subEntityNames.add("CLI_NON_HM_L17");
        subEntityNames.add("CLI_NON_HM_WAGE_SLIP");
        subEntityNames.add("EMPLOY_BEN_PARTNER");
        subEntityNames.add("PAR_NON_HM_L17");
        subEntityNames.add("PAR_NON_HM_WAGE_SLIP");
        subEntityNames.add("PAR_EMPLOY_KNOWN_CHANGE");

/* Count number of lines of sub-entities in array
        int cnt = 0;
        for (int d=0; d < value.length-1; d++) {
            if (subEntityNames.contains(value[d][2])) {
                cnt++;
            }
        }

// Create 2D array to hold sub-entities
        if (cnt > 0) {
            String[][] subEntities = new String[cnt][value[1].length];
            for (int e = 0; e < cnt; e++) {
                if (subEntityNames.contains(value[d][2])) {
                    subEntities[e] = value[e];
                }
            }
        }
*/
        System.out.println(Arrays.deepToString(value).replace("], ", "]\n").replace("[[", "[").replace("]]", "]"));
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
                if (listStrings.contains(value[b][2])) {
                    String entityName = value[b][2];
                    String inst = "";
                    inst = value[b][3];
                    if (!value[b - 1][2].equals(entityName)) {
                        writer.println("<entity id=" + "\"" + value[b][2] + "\"" + ">");
                    }
                    if (!value[b][3].equals(value[b - 1][3])) {
                        writer.println("<instance id=" + "\"" + value[b][3] + "\"" + ">");
                    }
                    if (value[b][3].equals(inst)) {
                        writer.println("<attribute id=" + "\"" + value[b][4] + "\"" + ">");
                        if (value[b][5].equals("~ ~") ||
                                ((value[b][4].contains("_D_") || value[b][4].contains("DATE") || value[b][4].contains("DOB"))
                                        && value[b][5].contains("BLANK"))) {
                            writer.println("<uncertain-val/>");
                        } else if ((value[b][5].equals("0") || value[b][5].contains("BLANK")) && value[b][4].contains("_T_")) {
                            writer.println("<text-val/>");
 //                       } else if (value[b][4].contains("_T_") ||value[b][4].equals("BANKACC_INPUT_N_7WP2_5A") || value[4][b].equals("LEVEL_OF_SERVICE")) {
                        } else if (value[b][4].contains("_T_")||value[b][4].equals("BANKACC_INPUT_N_7WP2_5A")|| value[4][b].equals("LEVEL_OF_SERVICE")) {
                            writer.println("<text-val>" + value[b][5] + "</text-val>");
                        } /*else if (value[b][4].contains("_B_") || value[b][5].equals("true") || value[b][5].equals("false")) {
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
                        }*/
                        writer.println("</attribute>");
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
