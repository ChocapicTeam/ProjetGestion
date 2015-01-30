public class UE {
	private String nom;

	public static enum types {
		IMPOSEE, CHOIX, VALIDEE, REDOUBLANT
	};

	private types type;

	public UE(types type, String nom) {
		this.type = type;
		this.nom = nom;
	}

	public types getType() {
		return type;
	}

	public void setType(types type) {
		this.type = type;
	}

	public String getNom() {
		return nom;
	}

	public void setNom(String nom) {
		this.nom = nom;
	}

	
	//Retourne le nom de l'UE à partir d'une chaine de caractère spécifique 
	public static String parseNomUE(String chaine) {
		String result = chaine;
		result = result.substring(result.indexOf("[") + 1);
		result = result.substring(0, result.indexOf("]"));
		return result;
	}
	
	
	//Retourne le type de l'UE à partir d'une chaine de caractère spécifique 
	public static types getType(String chaine) {
		if (chaine.toLowerCase().contains("choix"))
			return types.CHOIX;
		else if (chaine.toLowerCase().contains("impo")
				|| chaine.toLowerCase().contains("commun"))
			return types.IMPOSEE;
		else if (chaine.toLowerCase().contains("redou"))
			return types.REDOUBLANT;
		return types.CHOIX; // ne devrait jamais arriver
	}

	//Methode d'affichage de l'UE 
	public String toString() {
		return (" UE : " + nom + " type : " + type);
	}
}
