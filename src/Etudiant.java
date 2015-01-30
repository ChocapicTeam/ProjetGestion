import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;
import java.util.ArrayList;

public class Etudiant implements Serializable {
	private String numeroEtudiant;
	private Specialite specialite;
//kjnfslknfdskjhfdslkfndskjfsdlkfnsdkjfnsdkfbdskjnfdgkjbdskj
	private String prenom;
	private String nom;
	private ArrayList<UE> listeUE;
	private String mailPerso;
	private boolean isRedoublant;

	public Etudiant() {
		listeUE = new ArrayList<UE>();
		isRedoublant = false;
	}

	public String getNumero() {
		return numeroEtudiant;
	}

	public void setNumero(String numero) {
		this.numeroEtudiant = numero;
	}

	public String getPrenom() {
		return prenom;
	}

	public void setPrenom(String prenom) {
		this.prenom = prenom;
	}

	public Specialite getSpecialite() {
		return specialite;
	}

	public void setSpecialite(Specialite specialite) {
		this.specialite = specialite;
	}

	public String getNom() {
		return nom;
	}

	public void setNom(String nom) {
		this.nom = nom;
	}

	public String getMailPerso() {
		return mailPerso;
	}

	public void setMailPerso(String mailPerso) {
		this.mailPerso = mailPerso;
	}

	public boolean isRedoublant() {
		return isRedoublant;
	}

	public void setRedoublant(boolean isRedoublant) {
		this.isRedoublant = isRedoublant;
	}

	public String toString() {
		String result = nom + " " + prenom + " " + specialite.getNom() + " "
				+ mailPerso + " " + isRedoublant + " " + numeroEtudiant + " ";
		for (int i = 0; i < listeUE.size(); i++) {
			result += listeUE.get(i) + " ";
		}

		return result;
	}

	public ArrayList<UE> getListeUE() {
		return listeUE;
	}

	public void setListeUE(ArrayList<UE> listeUE) {
		this.listeUE = listeUE;
	}

	
	public void sauvegarder(File file) {
		ObjectOutputStream oos = null;
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(file);
			oos = new ObjectOutputStream(fos);
			oos.writeObject(this);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {
				oos.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}

	public static Etudiant charger(File file) {
		ObjectInputStream ois = null;
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(file);
			ois = new ObjectInputStream(fis);
			return ((Etudiant) ois.readObject());
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {
				ois.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return null;
	}

}
