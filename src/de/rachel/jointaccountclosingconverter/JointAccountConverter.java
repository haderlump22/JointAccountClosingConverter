package de.rachel.jointaccountclosingconverter;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jopendocument.dom.spreadsheet.Cell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

public class JointAccountConverter {
        private Map<String, Integer> foundCoordinates = new HashMap<String, Integer>();
        private Map<String, String> sumOverviewDetails = new HashMap<String, String>();
        private record closingSumRowValues(String sumType, Integer idOfSummand) {};
        private List<closingSumRowValues> closingSumRowValues = new ArrayList<>();
        // for the Column Position we need when determine from Formula Values
        private String alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        StringBuilder contentBuffer = new StringBuilder();

    JointAccountConverter() throws IOException {
        /*
         * neet to know
         * getCellAt(column, row) !!!
         */

        File file = new File("Aufstellungen2017.ods");
        String outputFile = "ha_abschlusssummenImport.txt";
        SpreadSheet spreadSheet;
        spreadSheet = SpreadSheet.createFromFile(file);

        int anzahl = spreadSheet.getSheetCount();
        for (int i = 0; i < anzahl; i++) {
            if (spreadSheet.getSheet(i).getName().startsWith("Pivot")) {
                Sheet actualSheet = spreadSheet.getSheet(i);
                System.out.println(actualSheet.getName());
                findFirstCell(actualSheet);
                collectSumOverviewDetails(actualSheet);

                // {unplanned+==[.D3], planned-==[.D5]+[.D6]+[.D9], unplanned-==[.D4]+[.D2], planned+==[.D10]+[.D11]}

                for (String sumType : sumOverviewDetails.keySet()) {
                    System.out.println(sumOverviewDetails.get(sumType));
                    generateClosingSumRowValues(actualSheet, sumType, sumOverviewDetails.get(sumType));
                }

                for (closingSumRowValues dataRow : closingSumRowValues) {
                    contentBuffer.append("('" + dataRow.sumType + "', "+ dataRow.idOfSummand +"),\n");
                }

            };
        }



        // remove all from the last commata to the end of content
        contentBuffer = contentBuffer.delete(contentBuffer.length() - 2, contentBuffer.length());

        Files.writeString(Paths.get(outputFile), contentBuffer, StandardCharsets.UTF_8);



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

    private void collectSumOverviewDetails(Sheet actualSheet) {

        /// so is our Spread to read, and we start at coordinates where the value "sum-"
        /// is present
        /// and read the formulas that form the values
        /// |       |planned |unplanned
        /// |-------|--------|--------
        /// | sum-  | -202,3 | -603
        /// | sum+  | 50     | 160

        int startColumn = foundCoordinates.get("column");
        int startRow = foundCoordinates.get("row");

        // get the Value Formula if the Value not equal 0
        if (Float.valueOf(actualSheet.getCellAt(startColumn + 1, startRow).getValue().toString()) != 0)
            sumOverviewDetails.put("planned-",actualSheet.getCellAt(startColumn + 1, startRow).getFormula());
        if (Float.valueOf(actualSheet.getCellAt(startColumn + 1, startRow + 1).getValue().toString()) != 0)
            sumOverviewDetails.put("planned+",actualSheet.getCellAt(startColumn + 1, startRow + 1).getFormula());
        if (Float.valueOf(actualSheet.getCellAt(startColumn + 2, startRow).getValue().toString()) != 0)
            sumOverviewDetails.put("unplanned-",actualSheet.getCellAt(startColumn + 2, startRow).getFormula());
        if (Float.valueOf(actualSheet.getCellAt(startColumn + 2, startRow + 1).getValue().toString()) != 0)
            sumOverviewDetails.put("unplanned+",actualSheet.getCellAt(startColumn + 2, startRow + 1).getFormula());
    }

    private void generateClosingSumRowValues(Sheet actualSheet, String sumType, String formula) {

        // first we neet to split the formula that looks like this => "=[.D3]" or like
        // this => "=[.D5]+[.D6]+[.D9]"
        if (formula.contains("+")) {
            String toSplit = formula.replaceAll("=|\\[\\.|\\]", "");
            String[] cellIds = toSplit.split("\\+");

            for (String cellId : cellIds) {
                closingSumRowValues.add(new closingSumRowValues(sumType, Integer.valueOf(actualSheet
                        .getCellAt(alphabet.indexOf(cellId.charAt(0)) + 1, Integer.valueOf(cellId.substring(1)) - 1)
                        .getValue().toString())));
            }

        } else {
            String cellId = formula.replaceAll("=|\\[\\.|\\]", "");
            closingSumRowValues.add(new closingSumRowValues(sumType, Integer.valueOf(actualSheet
                        .getCellAt(alphabet.indexOf(cellId.charAt(0)) + 1, Integer.valueOf(cellId.substring(1)) - 1)
                        .getValue().toString())));
        }
    }
}
