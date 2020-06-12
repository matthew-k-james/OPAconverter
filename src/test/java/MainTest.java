import org.junit.jupiter.api.Test;

import static org.junit.Assert.assertEquals;

public class MainTest {


    String[][] fileRow;

    public MainTest() {
        fileRow = new String[][]{
                {"null", "meansAssessment_PREPOP",  "BANKACC",  "the bank account1", "BANKACC_INPUT_B_7WP2_16A", "false"},
                {"null", "meansAssessment_PREPOP",  "BANKACC",  "the bank account1", "BANKACC_INPUT_T_7WP2_4A", "Halifax"},
                {"null", "meansAssessment_PREPOP",  "BANKACC",  "the bank account1", "BANKACC_INPUT_C_7WP2_18A", "30.00"},
                {"null", "meansAssessment_PREPOP",  "GLOBAL",  "global", "GB_INPUT_N_12WP3_1A", "60"},
                {"null", "meansAssessment_PREPOP",  "CARS_AND_MOTOR_VEHICLES",  "the cars &amp; motor vehicle1", "CARANDVEH_INPUT_D_14WP2_27A", "01-11-2018"},
                {"null", "meansAssessment_PREPOP",  "GLOBAL", "global", "BEN_AWARD_DATE", "~ ~"},
                {"null", "meansAssessment_PREPOP",  "GLOBAL", "global", "GB_DECL_T_38WP3_100A", "0"}
        };
    }

    @Test
    public void checkBooleanAttrTypeIsOutput() throws Exception {
        Main tester = new Main();

        String MeansDataModel = "/Users/matthewjames/Downloads/Means 18A Data Model.xlsx";
        String[][] dataModel = Main.readExcel(MeansDataModel);

        String result = Main.xdsDataStructure(dataModel, fileRow, 0);
        String expected = "<boolean-val>false</boolean-val>";
        //assert statements
        System.out.println("result is " + result + " expected result is " + expected);
        assertEquals(result,expected);
    }

    @Test
    public void checkTextAttrTypeIsOutput() throws Exception {
        Main tester = new Main();

        String MeansDataModel = "/Users/matthewjames/Downloads/Means 18A Data Model.xlsx";
        String[][] dataModel = Main.readExcel(MeansDataModel);

        String result = Main.xdsDataStructure(dataModel, fileRow, 1);
        String expected = "<text-val>Halifax</text-val>";
        //assert statements
        System.out.println("result is " + result + " expected result is " + expected);
        assertEquals(result,expected);
    }

    @Test
    public void checkCurrencyAttrTypeIsOutput() throws Exception {
        Main tester = new Main();

        String MeansDataModel = "/Users/matthewjames/Downloads/Means 18A Data Model.xlsx";
        String[][] dataModel = Main.readExcel(MeansDataModel);

        String result = Main.xdsDataStructure(dataModel, fileRow, 2);
        String expected = "<currency-val>30.00</currency-val>";
        //assert statements
        System.out.println("result is " + result + " expected result is " + expected);
        assertEquals(result,expected);
    }

    @Test
    public void checkNumberAttrTypeIsOutput() throws Exception {
        Main tester = new Main();

        String MeansDataModel = "/Users/matthewjames/Downloads/Means 18A Data Model.xlsx";
        String[][] dataModel = Main.readExcel(MeansDataModel);

        String result = Main.xdsDataStructure(dataModel, fileRow, 3);
        String expected = "<number-val>60</number-val>";
        //assert statements
        System.out.println("result is " + result + " expected result is " + expected);
        assertEquals(result,expected);
    }

    @Test
    public void checkDateAttrTypeIsOutput() throws Exception {
        Main tester = new Main();

        String MeansDataModel = "/Users/matthewjames/Downloads/Means 18A Data Model.xlsx";
        String[][] dataModel = Main.readExcel(MeansDataModel);

        String result = Main.xdsDataStructure(dataModel, fileRow, 4);
        String expected = "<date-val>2018-11-01</date-val>";
        //assert statements
        System.out.println("result is " + result + " expected result is " + expected);
        assertEquals(result,expected);
    }

    @Test
    public void checkUncertainAttrTypeIsOutput() throws Exception {
        Main tester = new Main();

        String MeansDataModel = "/Users/matthewjames/Downloads/Means 18A Data Model.xlsx";
        String[][] dataModel = Main.readExcel(MeansDataModel);

        String result = Main.xdsDataStructure(dataModel, fileRow, 5);
        String expected = "<uncertain-val/>";
        //assert statements
        System.out.println("result is " + result + " expected result is " + expected);
        assertEquals(result,expected);
    }

    @Test
    public void checkEscapedTextAttrTypeIsOutput() throws Exception {
        Main tester = new Main();

        String MeansDataModel = "/Users/matthewjames/Downloads/Means 18A Data Model.xlsx";
        String[][] dataModel = Main.readExcel(MeansDataModel);

        String result = Main.xdsDataStructure(dataModel, fileRow, 6);
        String expected = "<text-val/>";
        //assert statements
        System.out.println("result is " + result + " expected result is " + expected);
        assertEquals(result,expected);
    }

}