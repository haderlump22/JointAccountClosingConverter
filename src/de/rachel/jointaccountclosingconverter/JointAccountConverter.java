package de.rachel.jointaccountclosingconverter;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Date;
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

    private record closingSumRowValues(String sumType, Integer idOfSummand) {
    };

    private record closingDetailTableData(Integer abschlussDetailId, String kategorieBezeichnung,
            Float summeBetraege, Float planBetrag, Float differenz, String abschlussMonat, String bemerkung) {
    };

    private List<closingSumRowValues> closingSumRowValues = new ArrayList<>();
    private List<closingDetailTableData> closingDetailTableData = new ArrayList<>();
    // for the Column Position we need when determine from Formula Values
    private String alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    private StringBuilder contentBuffer = new StringBuilder();
    private int idForClosingDetailTable = 0;

    JointAccountConverter() throws IOException {
        File file = new File("Aufstellungen2017.ods");
        String outputFile = "ha_abschlusssummen.txt";
        String outputFileHaDetails = "ha_abschlussdetails.txt";
        SpreadSheet spreadSheet;
        spreadSheet = SpreadSheet.createFromFile(file);

        int anzahl = spreadSheet.getSheetCount();
        for (int i = 0; i < anzahl; i++) {
            if (spreadSheet.getSheet(i).getName().startsWith("Pivot")) {
                Sheet actualSheet = spreadSheet.getSheet(i);
                 System.out.println(actualSheet.getName());

                // create IDs for detailTable and Sum Values Data that need them
                createIdsForAccountClosingDetails(actualSheet);

                // block for Creating data for the Accountclosingdetail Table for Import
                collectDetailTableData(actualSheet);

                // Block for creating closingSum Data for Import
                findFirstCellOfSumArea(actualSheet);
                collectSumOverviewDetails(actualSheet);


                for (String sumType : sumOverviewDetails.keySet()) {
                    generateClosingSumRowValues(actualSheet, sumType, sumOverviewDetails.get(sumType));
                }

                // Block for creating Content for Import Files
                for (closingSumRowValues dataRow : closingSumRowValues) {
                    contentBuffer.append("('" + dataRow.sumType + "', "+ dataRow.idOfSummand +"),\n");
                }
            };
        }

        // save all changes
        spreadSheet.saveAs(file);

        // remove all from the last commata to the end of content
        contentBuffer = contentBuffer.delete(contentBuffer.length() - 2, contentBuffer.length());

        Files.writeString(Paths.get(outputFile), contentBuffer, StandardCharsets.UTF_8);



        // Sheet actualSheet = spreadSheet.getSheet(3);
        // if (!actualSheet.getCellAt("E1").getValue().equals("")){
        //     System.out.println("nich leer");
        // } else {
        //     System.out.println("leer");
        // }

        // spreadSheet.saveAs(file);

    }

    private void collectDetailTableData(Sheet actualSheet) {
        // run reading Information from Cell A1 to E<last Line before in Column A value
        // is "Gesamt ergebnis">
        Integer abschlussDetailId;
        String kategorieBezeichnung;
        Float summeBetraege;
        Float planBetrag;
        Float differenz;
        String abschlussMonat;
        String bemerkung;

        abschlussMonat = "01." + actualSheet.getName().replaceAll("Pivot-Tabelle_|_\\d{1,2}", "").replace("-", ".");

        if (actualSheet.getCellAt("A1").getValue().equals("Art")) {
            // we start in Row 2 (base that first row has index 0)
            int i = 1;
            while (!actualSheet.getCellAt(0, i).getValue().equals("Gesamt Ergebnis")) {
                abschlussDetailId = Integer.valueOf(actualSheet.getCellAt(19, i).getValue().toString());
                kategorieBezeichnung = actualSheet.getCellAt(0, i).getValue().toString();
                summeBetraege = Float.valueOf(actualSheet.getCellAt(1, i).getValue().toString());
                planBetrag = Float.valueOf(actualSheet.getCellAt(2, i).getValue().toString());
                differenz = Float.valueOf(actualSheet.getCellAt(3, i).getValue().toString());
                bemerkung = actualSheet.getCellAt(4, i).getValue().toString();

                closingDetailTableData.add(new closingDetailTableData(abschlussDetailId, kategorieBezeichnung,
                        summeBetraege, planBetrag, differenz, abschlussMonat, bemerkung));

                i++;
            }
        } else {
            System.err.println("Fehler... ID Feld in Zelle E1 in Sheet " + actualSheet.getName()
                    + " kann nicht gesetzt werden, schon ein Wert vorhanden!");
            System.exit(1);
        }
    }
    public static void main(String[] args) throws Exception {
        new JointAccountConverter();
    }

    private Map<String, Integer> findFirstCellOfSumArea(Sheet actualSheet) {
        // running fom Column 4 Row 0 to maximum column 7 row 30
        // to find the first cell with Value "Summe-",
        // so we can determine where the sumoverview Details are
        Cell<SpreadSheet> actualCell = null;

        // ensure we can work with rows and columns we want
        actualSheet.ensureRowCount(30);

        for (int myColumn = 4; myColumn < 8; myColumn++){
            for (int myRow = 0; myRow < 30; myRow++){
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
        if (Float.valueOf(actualSheet.getCellAt(startColumn + 1, startRow).getValue().toString()) != 0) {
            sumOverviewDetails.put("planned-",actualSheet.getCellAt(startColumn + 1, startRow).getFormula());
        } else {
            sumOverviewDetails.remove("planned-");
        }

        if (Float.valueOf(actualSheet.getCellAt(startColumn + 1, startRow + 1).getValue().toString()) != 0){
            sumOverviewDetails.put("planned+",actualSheet.getCellAt(startColumn + 1, startRow + 1).getFormula());
        } else {
            sumOverviewDetails.remove("planned+");
        }

        if (Float.valueOf(actualSheet.getCellAt(startColumn + 2, startRow).getValue().toString()) != 0){
            sumOverviewDetails.put("unplanned-",actualSheet.getCellAt(startColumn + 2, startRow).getFormula());
        } else {
            sumOverviewDetails.remove("unplanned-");
        }

        if (Float.valueOf(actualSheet.getCellAt(startColumn + 2, startRow + 1).getValue().toString()) != 0){
            sumOverviewDetails.put("unplanned+",actualSheet.getCellAt(startColumn + 2, startRow + 1).getFormula());
        } else {
            sumOverviewDetails.remove("unplanned+");
        }

    }

    private void generateClosingSumRowValues(Sheet actualSheet, String sumType, String formula) {

        // first we neet to split the formula that looks like this => "=[.D3]" or like
        // this => "=[.D5]+[.D6]+[.D9]"
        // or like this => "=SUM([.D6];[.D7];[.D11];[.D13])"
        if (formula.contains("+") || formula.contains("=SUM")) {
            if (formula.startsWith("=SUM")) {
                String toSplit = formula.replaceAll("=SUM|\\(|\\[|\\.|\\]|\\)", "");
                String[] cellIds = toSplit.split("\\;");

                for (String cellId : cellIds) {
                    // System.out.println("CellIdToGenerateId: " + cellId);
                    closingSumRowValues.add(new closingSumRowValues(sumType, Integer.valueOf(actualSheet
                            .getCellAt(19, Integer.valueOf(cellId.substring(1)) - 1)
                            .getValue().toString())));
                }
            } else {
                String toSplit = formula.replaceAll("=|\\[\\.|\\]", "");
                String[] cellIds = toSplit.split("\\+");

                for (String cellId : cellIds) {
                    // System.out.println("CellIdToGenerateId: " + cellId);
                    closingSumRowValues.add(new closingSumRowValues(sumType, Integer.valueOf(actualSheet
                            .getCellAt(19, Integer.valueOf(cellId.substring(1)) - 1)
                            .getValue().toString())));
                }
            }
        } else {
            String cellId = formula.replaceAll("=|\\[\\.|\\]", "");
            // System.out.println("CellIdToGenerateId: " + cellId);
            closingSumRowValues.add(new closingSumRowValues(sumType, Integer.valueOf(actualSheet
                        .getCellAt(19, Integer.valueOf(cellId.substring(1)) - 1)
                        .getValue().toString())));
        }
    }

    private void createIdsForAccountClosingDetails(Sheet actualSheet) {
        // we start only if there are nothin in E1
        // and we write the ID at a column that do not interupt the process
        // anywhere at column T or later

        // ensure we can work with Column until T
        actualSheet.ensureColumnCount(20);

        if (actualSheet.getImmutableCellAt(19, 0).getValue().equals("")){
            actualSheet.getCellAt(19, 0).setValue("ID");
            int i = 1;
            while (!actualSheet.getCellAt(0, i).getValue().equals("Gesamt Ergebnis")) {
                actualSheet.getCellAt(19, i).setValue(idForClosingDetailTable);
                idForClosingDetailTable++;
                i++;
            }
        } else {
            // We assume that only "ID" means that everything is prepared
            if (!actualSheet.getCellAt("T1").getValue().equals("ID")) {
                System.err.println("Fehler... ID Feld in Zelle E1 in Sheet "+ actualSheet.getName() + " kann nicht gesetzt werden, schon ein Wert vorhanden!");
                System.exit(1);
            }
        }

    }
}
