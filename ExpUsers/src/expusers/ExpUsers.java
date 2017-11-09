package expusers;

import bkgpi2a.AgencyUserQueryView;
import bkgpi2a.CallCenterUser;
import bkgpi2a.ClientAccountManager;
import bkgpi2a.Executive;
import bkgpi2a.PatrimonyManager;
import bkgpi2a.PatrimonyUserQueryView;
import bkgpi2a.SuperUser;
import bkgpi2a.User;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.mongodb.BasicDBObject;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PaperSize;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

/**
 * Programme pour exporter les utilisateurs d'un site Web dans un fichier Excel
 *
 * @author Thierry Baribaud
 * @version 0.10
 */
public class ExpUsers {

    private final static String path = "c:\\temp";

    private final static String filename = "users.xlsx";

    private final static String HOST = "10.65.62.133";
    private final static int PORT = 27017;

    /**
     * @param args arguments en ligne de commande
     */
    public static void main(String[] args) {

        FileOutputStream out;
        XSSFWorkbook classeur;
        XSSFSheet feuille;
        XSSFRow titre;
        XSSFCell cell;
        XSSFRow ligne;
        XSSFCellStyle cellStyle;
        XSSFCellStyle titleStyle;
        XSSFCellStyle cellStyle2;
        ObjectMapper objectMapper;
        User user;
        MongoDatabase mongoDatabase;
        MongoClient MyMongoClient;
        List<AgencyUserQueryView> managedAgencies;
        StringBuffer agencyList;
        List<PatrimonyUserQueryView> managedPatrimonies;
        StringBuffer patrimonyList;

        objectMapper = new ObjectMapper();

        MyMongoClient = new MongoClient(HOST, PORT);
        mongoDatabase = MyMongoClient.getDatabase("extranet-dev");

        MongoCollection<Document> MyCollection = mongoDatabase.getCollection("users");
        System.out.println(MyCollection.count() + " utilisateurs");

//      Création d'un classeur Excel
        classeur = new XSSFWorkbook();
        feuille = classeur.createSheet("Utilisateurs");

        // Style de cellule avec bordure noire
        cellStyle = classeur.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

        // Style pour le titre
        titleStyle = (XSSFCellStyle) cellStyle.clone();
        titleStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        titleStyle.setFillPattern(FillPatternType.LESS_DOTS);
//        titleStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());

        // Style pour les cellules à renvoi à la ligne automatique
        cellStyle2 = (XSSFCellStyle) cellStyle.clone();
        cellStyle2.setAlignment(HorizontalAlignment.JUSTIFY);
        cellStyle2.setVerticalAlignment(VerticalAlignment.JUSTIFY);
      
        // Ligne de titre
        titre = feuille.createRow(0);
        cell = titre.createCell((short) 0);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Nom");
        cell = titre.createCell((short) 1);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Prénom");
        cell = titre.createCell((short) 2);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Niveau");
        cell = titre.createCell((short) 3);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Mail");
        cell = titre.createCell((short) 4);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Etat");
        cell = titre.createCell((short) 5);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Société");
        cell = titre.createCell((short) 6);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Agences supervisées");
        cell = titre.createCell((short) 7);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("patrimoine supervisé");

        // Lit les ustilisateurs classés par nom et prénom
        MongoCursor<Document> MyCursor
                = MyCollection.find().sort(new BasicDBObject("lastName", 1).append("firstName", 1)).iterator();
        int n = 0;
        try {
            while (MyCursor.hasNext()) {
                user = objectMapper.readValue(MyCursor.next().toJson(), User.class);
                System.out.println(n
                        + " lastName:" + user.getLastName()
                        + ", firstName:" + user.getFirstName()
                        + ", status:" + user.getIsActive()
                        + ", login:" + user.getLogin()
                        + ", class:" + user.getClass().getSimpleName());
                n++;
                ligne = feuille.createRow(n);

                cell = ligne.createCell(0);
                cell.setCellValue(user.getLastName());
                cell.setCellStyle(cellStyle);

                cell = ligne.createCell(1);
                cell.setCellValue(user.getFirstName());
                cell.setCellStyle(cellStyle);

                if (user instanceof Executive) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((Executive) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellValue(((Executive) user).getCompany().getLabel());
                    cell.setCellStyle(cellStyle);

                    managedAgencies = ((Executive) user).getManagedAgencies();
                    System.out.println("  Managed agencies : " + managedAgencies);

                    agencyList = new StringBuffer();
                    for (AgencyUserQueryView agency : managedAgencies) {
                        if (agencyList.length() > 0) {
                            agencyList.append(", " + agency.getLabel());
                        } else {
                            agencyList.append(agency.getLabel());
                        }
                    }
                    if (agencyList.length() > 0) {
                        cell = ligne.createCell(6);
                        cell.setCellValue(agencyList.toString());
                        cell.setCellStyle(cellStyle2);
                    }

                } else if (user instanceof PatrimonyManager) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((PatrimonyManager) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellValue(((PatrimonyManager) user).getCompany().getLabel());
                    cell.setCellStyle(cellStyle);

                    managedAgencies = ((PatrimonyManager) user).getManagedAgencies();
                    System.out.println("  Managed agencies : " + managedAgencies);

                    agencyList = new StringBuffer();
                    for (AgencyUserQueryView agency : managedAgencies) {
                        if (agencyList.length() > 0) {
                            agencyList.append(", " + agency.getLabel());
                        } else {
                            agencyList.append(agency.getLabel());
                        }
                    }
                    if (agencyList.length() > 0) {
                        cell = ligne.createCell(6);
                        cell.setCellValue(agencyList.toString());
                        cell.setCellStyle(cellStyle2);
                    }
                    
                    managedPatrimonies = ((PatrimonyManager) user).getManagedPatrimonies();
                    System.out.println("  Managed patrimonies : " + managedPatrimonies);

                    patrimonyList = new StringBuffer();
                    for (PatrimonyUserQueryView patrimony : managedPatrimonies) {
                        if (patrimonyList.length() > 0) {
                            patrimonyList.append(", " + patrimony.getRef());
                        } else {
                            patrimonyList.append(patrimony.getRef());
                        }
                    }
                    if (patrimonyList.length() > 0) {
                        cell = ligne.createCell(7);
                        cell.setCellValue(patrimonyList.toString());
                        cell.setCellStyle(cellStyle2);
                    }
                    
                } else if (user instanceof CallCenterUser) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((CallCenterUser) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellValue(((CallCenterUser) user).getCompany().getLabel());
                    cell.setCellStyle(cellStyle);
                } else if (user instanceof ClientAccountManager) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((ClientAccountManager) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellValue(((ClientAccountManager) user).getCompany().getLabel());
                    cell.setCellStyle(cellStyle);
                } else if (user instanceof SuperUser) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((SuperUser) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellStyle(cellStyle);
                }

                cell = ligne.createCell(3);
                cell.setCellValue(user.getLogin());
                cell.setCellStyle(cellStyle);

                cell = ligne.createCell(4);
                if (user.getIsActive()) {
                    cell.setCellValue("Actif");
                } else {
                    cell.setCellValue("Inactif");
                }
                cell.setCellStyle(cellStyle);
            }

            // Ajustement automatique de la largeur des 6 premières colonnes
            for (int k = 0; k < 6; k++) {
                feuille.autoSizeColumn(k);
            }

            // Largeur des deux dernières colonnes fixées à 50 = 12 800 / 256
            feuille.setColumnWidth((int)6, (int)12800);
            feuille.setColumnWidth((int)7, (int)12800);
            
            // Format A4 en sortie
            feuille.getPrintSetup().setPaperSize(PaperSize.A4_PAPER);

            // Orientation paysage
            feuille.getPrintSetup().setLandscape(true);

            // Ajustement à une page en largeur
            feuille.setFitToPage(true);
            feuille.getPrintSetup().setFitWidth((short) 1);
            feuille.getPrintSetup().setFitHeight((short) 0);

            // En-tête et pied de page
            Header header = feuille.getHeader();
            header.setLeft("Liste des utilisateurs Extranet Anstel");
            header.setRight("&F");

            Footer footer = feuille.getFooter();
            footer.setLeft("Documentation confidentielle Anstel");
            footer.setCenter("Page &P / &N");
            footer.setRight("&D");

            // Ligne à répéter en haut de page
            feuille.setRepeatingRows(CellRangeAddress.valueOf("1:1"));

        } catch (IOException ex) {
            Logger.getLogger(ExpUsers.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            MyCursor.close();
        }

        // Enregistrement du classeur dans un fichier
        try {
            out = new FileOutputStream(new File(path + "\\" + filename));
            classeur.write(out);
            out.close();
            System.out.println("Fichier Excel " + filename + " créé dans " + path);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExpUsers.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExpUsers.class.getName()).log(Level.SEVERE, null, ex);
        }

    }
}
