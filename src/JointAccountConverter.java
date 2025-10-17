import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jopendocument.dom.spreadsheet.Cell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

public class JointAccountConverter {

    JointAccountConverter() throws IOException {
        Map<String, Integer> foundCoordinates = new HashMap<String, Integer>();
        Map<String, String> sumOverviewDetails = new HashMap<String, String>();
        // for the Column Position we need when determine from Formula Values
        String alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        /*
         * neet to know
         * getCellAt(column, row) !!!
         */

         /*
          * map.put("name", "demo");
          * map.put("fname", "fdemo");
          */
        File file = new File("Aufstellungen2017.ods");
        SpreadSheet spreadSheet;

        spreadSheet = SpreadSheet.createFromFile(file);

        Sheet actualSheet = spreadSheet.getSheet(1);
        foundCoordinates = findFirstCell(actualSheet);

        sumOverviewDetails = collectSumOverviewDetails(actualSheet, foundCoordinates);



        // int anzahl = spreadSheet.getSheetCount();
        // for (int i = 0; i < anzahl; i++) {
        //     if (spreadSheet.getSheet(i).getName().startsWith("Pivot")) {
        //         Sheet actualSheet = spreadSheet.getSheet(i);

        //         System.out.println(actualSheet.getCellAt("G6").getValue());
        //     };
        // }

        // System.out.println(actualSheet.getCellAt(6,5).get);

        // System.out.println(foundCell.);

//         System.out.println(actualSheet.getName());
// //

    }
    public static void main(String[] args) throws Exception {
        new JointAccountConverter();
    }

    private Map<String, Integer> findFirstCell(Sheet actualSheet) {
        // running fom Column 4 Row 0 to maximum column 7 row 13
        // to find the first cell with Value "Summe-",
        // so we can determine where the sumoverview Details are
        Cell<SpreadSheet> actualCell = null;
        Map<String, Integer> foundCoordinates = new HashMap<String, Integer>();

        for (int myColumn = 4; myColumn < 8; myColumn++){
            for (int myRow = 0; myRow < 14; myRow++){
                actualCell = actualSheet.getCellAt(myColumn, myRow);

                if (actualCell.getValue().equals("Summe-")) {
                    foundCoordinates.put("column", myColumn);
                    foundCoordinates.put("row", myRow);
                }
            }
        }
        return foundCoordinates;
    }

    private Map<String, String> collectSumOverviewDetails(Sheet actualSheet, Map<String, Integer> foundCoordinates) {

        /// so is our Spread to read, and we start at coordinates where the value "sum-"
        /// is present
        /// and read the formulas that form the values
        /// |       |planned |unplanned
        /// |-------|--------|--------
        /// | sum-  | -202,3 | -603
        /// | sum+  | 50     | 160

        Map<String, String> sumOverviewDetails = new HashMap<String, String>();
        int startColumn = foundCoordinates.get("column");
        int startRow = foundCoordinates.get("row");

        // get the Value Formula
        sumOverviewDetails.put("planned-",actualSheet.getCellAt(startColumn + 1, startRow).getFormula());
        sumOverviewDetails.put("planned+",actualSheet.getCellAt(startColumn + 1, startRow + 1).getFormula());
        sumOverviewDetails.put("unplanned-",actualSheet.getCellAt(startColumn + 2, startRow).getFormula());
        sumOverviewDetails.put("unplanned+",actualSheet.getCellAt(startColumn + 2, startRow + 1).getFormula());

        return sumOverviewDetails;

    }
}
