import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class TestExcel {

	private static final String fic = "results.xlsx";
	private static final String feuille1 = "M1 Informatique - S8 - Pre-insc";
	private static final String ficSauvEtudiant = "etudiants";
	private static final int indEtudiantFin = 6, indDebutUE = 8;

	
	public static void readXLSWithPOI(String nameFile) throws InvalidFormatException,
			IOException {
		// TODO Auto-generated method stub

		//Fichier de sauvegarde des étudiants 
		File fileEtu = new File(ficSauvEtudiant);
		//Ouverture du fichier Excel 
		File file = new File(nameFile);
		Workbook wb = WorkbookFactory.create(file);
		//On recupère la feuille 1 du fichier Excel 
		Sheet sh = wb.getSheet(feuille1);

		int lastRowNum = sh.getLastRowNum() + 1;
		Row row = sh.getRow(0);


		ArrayList<String> lstIndEtudiant = new ArrayList<String>();
		ArrayList<UE> lstUE = new ArrayList<UE>();
		ArrayList<Etudiant> lstEtudiant = new ArrayList<Etudiant>();

		//Récupération de l'entete du fichier pour les attribut d'etudiants 
		for (int i = 0; i < indEtudiantFin; i++) {
			Cell cell = row.getCell(i);
			lstIndEtudiant.add(cell.getStringCellValue());
		}
		//Récupération de l'entete des UEs et parsing de la chaine ( type / nom de l'UE ) 
		for (int i = indDebutUE; i < row.getLastCellNum(); i++) {
			Cell cell = row.getCell(i);
			lstUE.add(new UE(UE.getType(cell.getStringCellValue()), UE
					.parseNomUE(cell.getStringCellValue())));
		}

		//Initialisation des objet Etudiants et ajout dans une liste 
		for (int i = 1; i < lastRowNum; i++) {
			row = sh.getRow(i);
			Etudiant e = new Etudiant();
			for (int j = 0; j < indEtudiantFin; j++) {
				Cell cell = row.getCell(j);
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					switch (lstIndEtudiant.get(j)) {
					case "non":
						e.setNom(cell.getStringCellValue());
						break;
					case "prenom":
						e.setPrenom(cell.getStringCellValue());
						break;
					case "mail":
						e.setMailPerso(cell.getStringCellValue());
						break;
					case "specialite":
						e.setSpecialite(new Specialite(cell
								.getStringCellValue()));
						break;
					case "redoublant":
						if (cell.getStringCellValue().trim().toLowerCase()
								.equals("y")) {
							e.setRedoublant(true);
						} else {
							e.setRedoublant(false);
						}

						break;
					}
					break;
				case Cell.CELL_TYPE_NUMERIC:
					if (lstIndEtudiant.get(j).equals("numetu"))
						e.setNumero(String.valueOf((int) cell
								.getNumericCellValue()));
					break;
				}

			}
			//e.sauvegarder(fileEtu);
			lstEtudiant.add(e);
		}

		
		//Initialiosation des liste des liste d'UEs pour chaque etudiants
		for (int i = 1; i < lastRowNum; i++) {
			row = sh.getRow(i);
			for (int j = indDebutUE; j < row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);

				if (cell == null)
					continue;

				switch (cell.getCellType()) {

				case Cell.CELL_TYPE_STRING:
					switch (cell.getStringCellValue().trim().toLowerCase()) {
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
					break;

				default:
					break;
				}

			}
		}
		
		//Affichage de la liste d'étudiants 
		for (int i = 0; i < lstEtudiant.size(); i++) {
			System.out.println(lstEtudiant.get(i).toString());
		}

	}

}
