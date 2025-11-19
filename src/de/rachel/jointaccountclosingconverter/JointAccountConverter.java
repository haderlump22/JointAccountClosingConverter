package de.rachel.jointaccountclosingconverter;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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

    private record closingSumBalanceAllocationValues(Integer partyId, Float percentOfShare, Float valueOfShare,
            String closingMonth, String postgresTimestampFunction, String closingComment) {
    };

    private List<closingSumRowValues> closingSumRowValues = new ArrayList<>();
    private List<closingDetailTableData> closingDetailTableData = new ArrayList<>();
    private List<closingSumBalanceAllocationValues> closingSumBalanceAllocationValues = new ArrayList<>();
    // for the Column Position we need when determine from Formula Values
    private StringBuilder contentBufferAbschlusssummen = new StringBuilder();
    private StringBuilder contentBufferAbschlussDetails = new StringBuilder();
    private StringBuilder contentBufferAbschlussAufteilung = new StringBuilder();
    private int idForClosingDetailTable = 0;

    JointAccountConverter() throws IOException {
        String[] year = {"2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025"};
        // String[] year = {"2017"};
        String outputFileHaSummen = "ha_abschlusssummen.txt";
        String outputFileHaDetails = "ha_abschlussdetails.txt";
        String outputFileBalanceAllocationShare = "ha_abschlusssummen_aufteilung.txt";
        String abschlussMonat;

        for (String actualYear : year) {
            File file = new File("Aufstellungen" + actualYear + ".ods");
            SpreadSheet spreadSheet = SpreadSheet.createFromFile(file);

            int anzahl = spreadSheet.getSheetCount();
            for (int i = 0; i < anzahl; i++) {
                if (spreadSheet.getSheet(i).getName().startsWith("Pivot")) {
                    Sheet actualSheet = spreadSheet.getSheet(i);
                    // if (spreadSheet.getSheet(i).getName().startsWith("Pivot-Tabelle_05-2023")) {
                    //     System.out.println("Breackpoint");
                    // }
                    System.out.println("processing: " + actualSheet.getName());

                    // set the current closingMonth
                    abschlussMonat = "01." + actualSheet.getName().replaceAll("Pivot-Tabelle_|_\\d{1,2}", "").replace("-", ".");

                    // create IDs for detailTable and Sum Values Data that need them
                    System.out.println("...create IDs");
                    createIdsForAccountClosingDetails(actualSheet);

                    // block for Creating data for the Accountclosingdetail Table for Import
                    System.out.println("...collect Detail Data");
                    collectDetailTableData(actualSheet, abschlussMonat);

                    // Block for creating closingSum Data for Import
                    System.out.println("...find SumOverview Block");
                    findFirstCellOfSumArea(actualSheet);
                    System.out.println("...collect SumOverview Details");
                    collectSumOverviewDetails(actualSheet, abschlussMonat);

                    System.out.println("...create SumOverview DataRows");
                    for (String sumType : sumOverviewDetails.keySet()) {
                        generateClosingSumRowValues(actualSheet, sumType, sumOverviewDetails.get(sumType));
                    }

                    // Block for creating Content for Import Files
                    System.out.println("...create SumOverview Import Data");
                    for (closingSumRowValues dataRow : closingSumRowValues) {
                        contentBufferAbschlusssummen.append("('" + dataRow.sumType + "', "+ dataRow.idOfSummand +"),\n");
                    }

                    // now we clean the list for the Values from the next Sheet
                    closingSumRowValues.clear();

                    System.out.println("...create Closingdetail Import Data");
                    for (closingDetailTableData dataRow : closingDetailTableData) {
                        contentBufferAbschlussDetails.append("(" + dataRow.abschlussDetailId +", '"+ dataRow.kategorieBezeichnung + "', " + dataRow.summeBetraege
                                        + ", " + dataRow.planBetrag + ", " + dataRow.differenz + ", '" + dataRow.abschlussMonat + "', '" + dataRow.bemerkung + "'),\n");
                    }

                    // now we clean the list for the Values from the next Sheet
                    closingDetailTableData.clear();

                    System.out.println("...create ClosingBalanceAllocation Import Data");
                    for (closingSumBalanceAllocationValues dataRow : closingSumBalanceAllocationValues) {
                        contentBufferAbschlussAufteilung.append("(" + dataRow.partyId +", "+ dataRow.percentOfShare + ", " + dataRow.valueOfShare
                                        + ", '" + dataRow.closingMonth + "', " + dataRow.postgresTimestampFunction + ", '" + dataRow.closingComment + "'),\n");
                    }

                    // now we clean the list for the Values from the next Sheet
                    closingSumBalanceAllocationValues.clear();

                };
            }

            // save all changes
            System.out.println("...save change ods File");
            spreadSheet.saveAs(file);
            System.out.println("LetzteID für die Summen des Jahres (" + actualYear + "): " + idForClosingDetailTable);



            // Sheet actualSheet = spreadSheet.getSheet(3);
            // if (!actualSheet.getCellAt("E1").getValue().equals("")){
            //     System.out.println("nich leer");
            // } else {
            //     System.out.println("leer");
            // }

            // spreadSheet.saveAs(file);
        }

        // remove all from the last commata to the end of content and write it to Importfile
        System.out.println("...write SumOverview Data to Importfile");
        contentBufferAbschlusssummen.delete(contentBufferAbschlusssummen.length() - 2, contentBufferAbschlusssummen.length());
        Files.writeString(Paths.get(outputFileHaSummen), contentBufferAbschlusssummen, StandardCharsets.UTF_8);

        System.out.println("...write Closingdetail Data to Importfile");
        contentBufferAbschlussDetails.delete(contentBufferAbschlussDetails.length() - 2, contentBufferAbschlussDetails.length());
        Files.writeString(Paths.get(outputFileHaDetails), contentBufferAbschlussDetails, StandardCharsets.UTF_8);

        System.out.println("...write BalanceAllocation Data to Importfile");
        contentBufferAbschlussAufteilung.delete(contentBufferAbschlussAufteilung.length() - 2, contentBufferAbschlussAufteilung.length());
        Files.writeString(Paths.get(outputFileBalanceAllocationShare), contentBufferAbschlussAufteilung, StandardCharsets.UTF_8);
    }

    private void collectDetailTableData(Sheet actualSheet, String abschlussMonat) {
        // run reading Information from Cell A1 to E<last Line before in Column A value
        // is "Gesamt ergebnis">
        Integer abschlussDetailId;
        String kategorieBezeichnung;
        Float summeBetraege = 0.0f;
        Float planBetrag = 0.0f;
        Float differenz = 0.0f;
        String bemerkung = "";



        if (actualSheet.getCellAt("A1").getValue().equals("Art") || actualSheet.getCellAt("A1").getValue().equals("Kategorie")) {
            // we start in Row 2 (base that first row has index 0)
            int i = 1;
            while (!actualSheet.getImmutableCellAt(0, i).getValue().equals("Gesamt Ergebnis")
                    && !actualSheet.getImmutableCellAt(0, i).getValue().equals("Summe Ergebnis")) {
                abschlussDetailId = Integer.valueOf(actualSheet.getImmutableCellAt(19, i).getValue().toString());
                kategorieBezeichnung = actualSheet.getImmutableCellAt(0, i).getValue().toString();
                if (!actualSheet.getImmutableCellAt(1, i).getValue().toString().isEmpty()) {
                    try {
                        summeBetraege = Float.valueOf(actualSheet.getImmutableCellAt(1, i).getValue().toString());
                    } catch (NumberFormatException e) {
                        System.out.println(e.getMessage());
                    }
                }

                if (!actualSheet.getImmutableCellAt(2, i).getValue().toString().isEmpty()) {
                    try {
                        planBetrag = Float.valueOf(actualSheet.getImmutableCellAt(2, i).getValue().toString());
                    } catch (NumberFormatException e) {
                        if (actualSheet.getImmutableCellAt(2, i).getValue().toString().matches("[ a-zA-Z0-9\\,üäö]*")) {
                            bemerkung = actualSheet.getImmutableCellAt(2, i).getValue().toString();
                        } else {
                            System.out.println(e.getMessage());
                        }
                    }
                }

                if (!actualSheet.getImmutableCellAt(3, i).getValue().toString().isEmpty()) {
                    try {
                        differenz = Float.valueOf(actualSheet.getImmutableCellAt(3, i).getValue().toString());
                    } catch (NumberFormatException e) {
                        System.out.println(e.getMessage());
                    }
                }

                // if value is not set before in the getting area for decimal Values
                // so we set it at this point with the Values where we expected the comment
                // for the detail Value
                if (bemerkung.isEmpty()) {
                    bemerkung = actualSheet.getImmutableCellAt(4, i).getValue().toString();
                }

                closingDetailTableData.add(new closingDetailTableData(abschlussDetailId, kategorieBezeichnung,
                        summeBetraege, planBetrag, differenz, abschlussMonat, bemerkung));

                i++;
            }
        } else {
            System.err.println("Fehler... die Pivot in " + actualSheet.getName()
                    + " hat keine Korrekte Feldbezeichungen!");
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

    private void collectSumOverviewDetails(Sheet actualSheet, String abschlussMonat) {
        StringBuilder comment = new StringBuilder();

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
        if (Float.valueOf(actualSheet.getImmutableCellAt(startColumn + 1, startRow).getValue().toString()) != 0) {
            sumOverviewDetails.put("planned-",actualSheet.getImmutableCellAt(startColumn + 1, startRow).getFormula());
        } else {
            sumOverviewDetails.remove("planned-");
        }

        if (Float.valueOf(actualSheet.getImmutableCellAt(startColumn + 1, startRow + 1).getValue().toString()) != 0){
            sumOverviewDetails.put("planned+",actualSheet.getImmutableCellAt(startColumn + 1, startRow + 1).getFormula());
        } else {
            sumOverviewDetails.remove("planned+");
        }

        if (Float.valueOf(actualSheet.getImmutableCellAt(startColumn + 2, startRow).getValue().toString()) != 0){
            sumOverviewDetails.put("unplanned-",actualSheet.getImmutableCellAt(startColumn + 2, startRow).getFormula());
        } else {
            sumOverviewDetails.remove("unplanned-");
        }

        if (Float.valueOf(actualSheet.getImmutableCellAt(startColumn + 2, startRow + 1).getValue().toString()) != 0){
            sumOverviewDetails.put("unplanned+",actualSheet.getImmutableCellAt(startColumn + 2, startRow + 1).getFormula());
        } else {
            sumOverviewDetails.remove("unplanned+");
        }

        /**
         * And with this coordinates we can collect at this Moment the
         * Balance Allocation Shares of this Month
         * after all 4 lines maybe comments
         */
        /// |       |        | 54,65%  | 45,35%  | comment?
        /// |-------|--------|---------|---------|
        /// |geplant| 117,40 | 64,16   | 53,24   | comment?
        /// |zusaetz| -351,80| -192,26 | -159,54 | comment?
        ///                  |---------|---------|
        ///                  |-128,10  |-106,30  | comment?

        // we collect all theoretical 4 comments
        for (int i = 0; i < 4; i++) {
            if (!actualSheet.getImmutableCellAt(startColumn + 4, startRow + (i+4)).getValue().toString().isEmpty()) {
                comment.append(actualSheet.getImmutableCellAt(startColumn + 4, startRow + (i+4)).getValue().toString() + "\n");
            }
        }

        // delete the last newline if there content exist
        if (comment.length() > 0) comment.delete(comment.length() - 1, comment.length());


        // for the first Person => in every Overviews ever me
        closingSumBalanceAllocationValues.add(new closingSumBalanceAllocationValues(
            2,
            Float.valueOf(actualSheet.getImmutableCellAt(startColumn + 2, startRow + 4).getValue().toString()),
            Float.valueOf(actualSheet.getImmutableCellAt(startColumn + 2, startRow + 7).getValue().toString()),
            abschlussMonat,
            "CURRENT_TIMESTAMP",
            "importiert von den ODS Tabellen zum Zeitpunkt" + (comment.length() > 0 ? "\n" : "") + comment));

        // and now for the second Person => in every Overviews my darling
        closingSumBalanceAllocationValues.add(new closingSumBalanceAllocationValues(
            6,
            Float.valueOf(actualSheet.getImmutableCellAt(startColumn + 3, startRow + 4).getValue().toString()),
            Float.valueOf(actualSheet.getImmutableCellAt(startColumn + 3, startRow + 7).getValue().toString()),
            abschlussMonat,
            "CURRENT_TIMESTAMP",
            "importiert von den ODS Tabellen zum Zeitpunkt" + (comment.length() > 0 ? "\n" : "") + comment));
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
                            .getImmutableCellAt(19, Integer.valueOf(cellId.substring(1)) - 1)
                            .getValue().toString())));
                }
            } else {
                String toSplit = formula.replaceAll("=|\\[\\.|\\]", "");
                String[] cellIds = toSplit.split("\\+");

                for (String cellId : cellIds) {
                    // System.out.println("CellIdToGenerateId: " + cellId);
                    closingSumRowValues.add(new closingSumRowValues(sumType, Integer.valueOf(actualSheet
                            .getImmutableCellAt(19, Integer.valueOf(cellId.substring(1)) - 1)
                            .getValue().toString())));
                }
            }
        } else {
            String cellId = formula.replaceAll("=|\\[\\.|\\]", "");
            // System.out.println("CellIdToGenerateId: " + cellId);
            closingSumRowValues.add(new closingSumRowValues(sumType, Integer.valueOf(actualSheet
                        .getImmutableCellAt(19, Integer.valueOf(cellId.substring(1)) - 1)
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
            // if there getting exists IDValues from alrady readed ods Files
            // we habe to inkrement it bevore setting its Value
            // for ods Sheets where they not exists

            actualSheet.getCellAt(19, 0).setValue("ID");
            int i = 1;
            while (!actualSheet.getImmutableCellAt(0, i).getValue().equals("Gesamt Ergebnis")
                    && !actualSheet.getImmutableCellAt(0, i).getValue().equals("Summe Ergebnis")) {
                actualSheet.getCellAt(19, i).setValue(idForClosingDetailTable);
                idForClosingDetailTable++;
                i++;
            }
        } else {
            // We assume that only "ID" means that everything is prepared
            if (!actualSheet.getImmutableCellAt("T1").getValue().equals("ID")) {
                System.err.println("Fehler... ID Feld in Zelle T1 in Sheet "+ actualSheet.getName() + " kann nicht gesetzt werden, schon ein Wert vorhanden!");
                System.exit(1);
            } else {
                // Whenever an ID is defined, we save it so that we can continue
                // with a subsequent ID for worksheets for which no ID is defined.
                for (int i = 1; !actualSheet.getImmutableCellAt(19, i).getValue().toString().isEmpty(); i++) {
                    idForClosingDetailTable = Integer.valueOf(actualSheet.getImmutableCellAt(19, i).getValue().toString());
                }

                // we increse the id for the next Sheet if there no one IDs is defines
                // so we can start to write withe correct next ID Value
                idForClosingDetailTable++;
            }
        }

    }
}
