package ch.next.sigvaris;

import java.util.Hashtable;
import java.util.Iterator;

import ch.next.sigvaris.model.Artikel;
import ch.next.sigvaris.model.Field;

import com.smartxls.WorkBook;

//Prototyp V1
public class MergeArtikelstamm {	

//	private final static String mode="Test";
	private final static String mode="";
	
	private  static String templateFilename = null;
	private  static String targetFilename = null;
	private static WorkBook targetwb = null;
	private static Hashtable<Integer, String> colsTarget = new Hashtable<Integer, String>();
	private static Hashtable<String, Field> fieldsTarget = new Hashtable<String, Field>();
	private static Hashtable<String, Artikel> artikel = new Hashtable<String, Artikel>();

	public static void main(String[] args) {
		System.out.println("Start");
		try {			
			templateFilename= "C:\\Sigvaris\\Sigvaris_Artikelstamm_Template"+mode+".xlsx";
			targetFilename = "C:\\Sigvaris\\Sigvaris_Artikelstamm_DACH"+mode+".xlsx";			
			targetwb = new WorkBook();
			targetwb.readXLSX(templateFilename);
			targetwb.setSheet(0);

			initTargetKeys(targetwb);
			
			read("C:\\Sigvaris\\DEAT"+mode+".xlsx");
			read("C:\\Sigvaris\\CH"+mode+".xlsx");
			read("C:\\Sigvaris\\FR"+mode+".xlsx");
			writeExcel(targetwb);

			targetwb.writeXLSX(targetFilename);
			targetwb.dispose();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		System.out.println("Finished");
	}

	private static void initTargetKeys(WorkBook wb) {
		try {
			int row = 0;
			int col = 0;
			// Keys des Templates einlesen
			String key = wb.getText(row, col).trim();
			while (key != null && !"".equals(key.trim())) {
				colsTarget.put(col, key);
				fieldsTarget.put(key, new Field(key, wb.getText(row+1, col).trim(), "string"));
				col++;
				key = wb.getText(row, col).trim();
			}
			return;
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	private static void read(String filename) {
		try {						
			System.out.println(filename);
			WorkBook wb = new WorkBook();
			System.out.println("Caching...");
			wb.readXLSX(filename);
			System.out.println("Reading...");
			wb.setSheet(0);

			int row = 3;
			int col = 0;

			// Inhalt einlesen im Prototyp zunächst alles als String vielleicht reicht das so			
			String artikelnummer = wb.getText(row, col).trim();
			while (artikelnummer != null && !"".equals(artikelnummer)) {
				System.out.println(row + ": " + artikelnummer);
				Artikel a = artikel.get(artikelnummer);
				if (a == null) {
					a = new Artikel();
				}
				a.Key = wb.getText(row, 0).trim();

				String colKey = wb.getText(0, col).trim();
				while (colKey != null && !"".equals(colKey)) {
					if (fieldsTarget.get(colKey) != null) {
						Field f = fieldsTarget.get(colKey);
						if ("double".equals(f.FieldType)) {
							a.doubleObjects.put(colKey, wb.getNumber(row, col));
						}
						if ("integer".equals(f.FieldType)) {
							a.integerObjects.put(colKey, Integer.valueOf(wb
									.getText(row, col).trim()));
						} else { // string
							String x = wb.getText(row, col)
									.trim();
							if (x != null && !"".equals(x)) {
								a.stringObjects.put(colKey, x);
							}
						}
						System.out.println(colKey + "/"
								+ fieldsTarget.get(colKey) + ": "
								+ wb.getText(row, col).trim());
					}
					col++;
					colKey = wb.getText(0, col).trim();
				}
				artikel.put(a.Key, a);
				row++;
				col = 0;
				artikelnummer = wb.getText(row, col).trim();
			}
			wb.dispose();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
			
	private static void writeExcel(WorkBook wb) {
		try {
			wb.setSheet(0);
			int row = 0;
			int col = 0;

			// Keys des Templates einlesen
			String key = wb.getText(row, col).trim();
			while (key != null && !"".equals(key)) {
				colsTarget.put(col, key);
				col++;
				key = wb.getText(row, col).trim();
			}

			col = 0;
			row = 2;
			// Daten schreiben
			Iterator<String> it = artikel.keySet().iterator();
			while (it.hasNext()) {
				String k = it.next();
				Artikel a = artikel.get(k);
				System.out.println(row + "/Writing: " + a.Key);
				Iterator<Integer> itCols = colsTarget.keySet().iterator();
				while (itCols.hasNext()) {
					col = itCols.next();
					String colkey = colsTarget.get(col);
					Field f = fieldsTarget.get(colkey);
					if (f != null) {
						if ("double".equals(f.FieldType)) {
							Double valD = a.doubleObjects.get(f.Key);
							if (valD != null) {
								wb.setNumber(row, col, valD);
							}
						}
						if ("integer".equals(f.FieldType)) {
							Integer valI = a.integerObjects.get(f.Key);
							if (valI != null) {
								wb.setNumber(row, col, valI);
							}
						} else { // string
							String valS = a.stringObjects.get(f.Key);
							if (valS != null) {
								wb.setText(row, col, valS);
							}
						}
					}
				}
				col = 0;
				row++;
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
}
