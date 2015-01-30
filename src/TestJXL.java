import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.*;




public class TestJXL {
	
	private static final String fic = "resultat.xls";
	private static final String feuille1 = "M1 Informatique - S8 - Pre-insc";
	private static final String ficSauvEtudiant = "etudiants";
	private static final int indEtudiantFin = 6, indDebutUE = 8;

	public static void main(String[] args) throws BiffException, IOException {
		Workbook workbook = Workbook.getWorkbook(new File("resultat.xls"));
		Sheet sheet1 = workbook.getSheet(0);

		int lastRowNum = sheet1.getRows();
		
		ArrayList<String> lstIndEtudiant = new ArrayList<String>();
		ArrayList<UE> lstUE = new ArrayList<UE>();
		ArrayList<Etudiant> lstEtudiant = new ArrayList<Etudiant>();

		
		for (int i = 0; i < indEtudiantFin; i++) {
			Cell cell = sheet1.getCell(i,0);
			lstIndEtudiant.add(cell.getContents());
		}
		
		for (int i = indDebutUE; i < sheet1.getColumns(); i++) {
			Cell cell = sheet1.getCell(i,0);
			lstUE.add(new UE(UE.getType(cell.getContents()), UE
					.parseNomUE(cell.getContents())));
		}
		
		for (int i = 1; i < lastRowNum; i++) {
			Etudiant e = new Etudiant();
			
			for (int j = 0; j < indEtudiantFin; j++) {
				Cell cell = sheet1.getCell(j,i);
				if (cell.getType() == CellType.LABEL) { 
					switch (lstIndEtudiant.get(j)) {
					case "non":
						e.setNom(cell.getContents());
						break;
					case "prenom":
						e.setPrenom(cell.getContents());
						break;
					case "mail":
						e.setMailPerso(cell.getContents());
						break;
					case "specialite":
						e.setSpecialite(new Specialite(cell
								.getContents()));
						break;
					case "redoublant":
						if (cell.getContents().trim().toLowerCase()
								.equals("y")) {
							e.setRedoublant(true);
						} else {
							e.setRedoublant(false);
						}
						break;
					}
				}
				else if (cell.getType() == CellType.NUMBER) {
					if (lstIndEtudiant.get(j).equals("numetu"))
						e.setNumero(String.valueOf( cell
								.getContents()));
				}

			}
			lstEtudiant.add(e);
		}
				
		
		for (int i = 1; i < lastRowNum; i++) {
			for (int j = indDebutUE; j < sheet1.getColumns(); j++) {
				Cell cell = sheet1.getCell(j,i);

				if (cell.getType() == CellType.EMPTY)
					continue;
				if (cell.getType() == CellType.LABEL) {
					switch (cell.getContents().trim().toLowerCase()) {
					case "y":
						if (!lstEtudiant.get(i - 1).isRedoublant())
							lstEtudiant.get(i - 1).getListeUE()
									.add(lstUE.get(j - indDebutUE));
						break;
					case "ad":
						(lstUE.get(j - indDebutUE)).setType(UE.types.VALIDEE);
						lstEtudiant.get(i - 1).getListeUE()
								.add(lstUE.get(j - indDebutUE));
						break;
					case "noad":
						(lstUE.get(j - indDebutUE)).setType(UE.types.CHOIX);
						lstEtudiant.get(i - 1).getListeUE()
								.add(lstUE.get(j - indDebutUE));
						break;
					default:
						break;
					}
				}

			}
		}
		
		for (int i = 0; i < lstEtudiant.size(); i++) {
			System.out.println(lstEtudiant.get(i).toString() + "\n");
		}
	}

}
