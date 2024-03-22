package ch.fablabwinti.accounting.main;

public class AccountConvertorPostFinanceCsvTest {

    public static void main(String[] args) {
        String text = "GUTSCHRIFT AUFTRAGGEBER: CROSS ING AG TECHNOPARKSTRASSE 2 CH 8406 WINTERT HUR MITTEILUNGEN: FABLAB ANNA TOPERT SPESENBETRAG 0.00 CHF SHA REFERENZEN: 6D580BBACA43F20BC7DADD5DB0421559 0156060DN8310043 240229CH0BFK0OM8";

        String t1 = text.substring(0, 25);
        String t2 = text.substring(25, 50);
        String t3 = text.substring(50, 75);
        String t4 = text.substring(75, 100);

        String tArr[] = {
                t1, t2, t3, t4
        };

        for (String t : tArr) {
            System.out.println(t);
        }
    }
}

